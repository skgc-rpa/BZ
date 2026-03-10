import os
import requests
import warnings
import urllib3
import pandas as pd
import smtplib
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime, timedelta
from io import StringIO
from io import BytesIO
from lxml import html
from email.message import EmailMessage

warnings.filterwarnings('ignore')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# 1. Session / Headers
# =========================================================
session = requests.Session()
session.verify = False  # 반드시 False

headers = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Content-Type': 'application/x-www-form-urlencoded',
    'Referer': 'https://www.ccfgroup.com/member/member.php',
}

BASE_URL = "https://www.ccfgroup.com"

# =========================================================
# 2. Login Function (session 반환)
# =========================================================
def login_ccfgroup(session, headers, login_data):
    """
    CCFGroup 로그인
    성공 시 로그인된 session 반환
    """
    login_url = "https://www.ccfgroup.com/member/member.php"

    resp = session.post(
        login_url,
        data=login_data,
        headers=headers,
        timeout=30
    )
    resp.raise_for_status()

    return session

# =========================================================
# 3. Daily / Weekly Finder
# =========================================================
today = datetime.today().date()
offset_days = 1
target_date = today - timedelta(days=offset_days)

def find_market_daily(list_url: str, title_prefix: str):
    resp = session.get(list_url, headers=headers, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    candidates = []

    for a in soup.find_all("a"):
        text = a.get_text(strip=True)
        if not text.startswith(title_prefix):
            continue

        try:
            date_str = text[text.find("(") + 1 : text.find(")")]
            post_date = datetime.strptime(date_str, "%b %d, %Y").date()
        except Exception:
            continue

        if post_date <= target_date:
            full_url = urljoin(BASE_URL, a.get("href"))
            candidates.append((post_date, full_url))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def find_market_weekly(list_url: str, title_prefix: str):
    resp = session.get(list_url, headers=headers, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    for a in soup.find_all("a"):
        if a.get_text(strip=True).startswith(title_prefix):
            return urljoin(BASE_URL, a.get("href"))

    return None

# =========================================================
# 4. URL Extract (비로그인)
# =========================================================
bz_daily = find_market_daily(
    "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=C00000",
    "Benzene market daily"
)

sm_daily = find_market_daily(
    "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=F00000",
    "Styrene monomer market daily"
)

bz_weekly = find_market_weekly(
    "https://www.ccfgroup.com/newscenter/index.php?Class_ID=200000&subclassid=C00000",
    "Benzene market weekly"
)

sm_weekly = find_market_weekly(
    "https://www.ccfgroup.com/newscenter/index.php?Class_ID=200000&subclassid=F00000",
    "Styrene monomer market weekly"
)

urls = {
    "bz_daily": bz_daily,
    "sm_daily": sm_daily,
    "bz_weekly": bz_weekly,
    "sm_weekly": sm_weekly
}

print("=== Extracted URLs (No Login) ===")
for k, v in urls.items():
    print(f"{k}: {v}")
df_url = pd.Series(urls).to_frame(name='URL')

# =========================================================
# 5. Login (GitHub Secrets 환경변수 적용)
# =========================================================
USERNAME = os.environ.get("CCF_USER")
PASSWORD = os.environ.get("CCF_PASSWORD")

if not USERNAME or not PASSWORD:
    print("⚠️ CCF_USER 또는 CCF_PASSWORD 환경변수가 누락되었습니다.")

login_data = {
    'custlogin': '1',
    'action': 'login',
    'username': USERNAME,
    'password': PASSWORD,
    'savecookie': 'savecookie'
}

session = login_ccfgroup(session, headers, login_data)
print("✅ 로그인 완료 (session 유지됨)")

# =========================================================
# 6. 로그인 상태로 URL 접근 → 테이블 추출
# =========================================================
def fetch_tables_as_df(session, url, headers):
    if not url:
        return []

    resp = session.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    
    html_file_like = StringIO(resp.text)
    dfs = pd.read_html(html_file_like)

    return dfs

# =========================================================
# 7. 사용 예 (데이터프레임 전처리)
# =========================================================
df_bz_daily = fetch_tables_as_df(session, bz_daily, headers)
df_bz_weekly = fetch_tables_as_df(session, bz_weekly, headers)
df_sm_daily = fetch_tables_as_df(session, sm_daily, headers)
df_sm_weekly = fetch_tables_as_df(session, sm_weekly, headers)

df_bz_daily_f = df_bz_daily[1].iloc[7:14,:].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_bz_weekly_f = df_bz_weekly[2].iloc[:13,:].T.drop_duplicates(keep='first').T.reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_bz_weekly_f_or = df_bz_weekly_f.iloc[:df_bz_weekly_f.iloc[:, 0].str.contains('Inventory', na=False).idxmax()].reset_index(drop=True)
df_bz_weekly_f_inv = df_bz_weekly_f.iloc[df_bz_weekly_f.iloc[:, 0].str.contains('Inventory', na=False).idxmax():].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_sm_weekly_f = df_sm_weekly[2].iloc[1:df_sm_weekly[2].iloc[:, 0].str.contains('Import & export', na=False).idxmax(), :].T.drop_duplicates(keep='first').T.reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_sm_weekly_f_or = df_sm_weekly_f.iloc[:df_sm_weekly_f.iloc[:, 0].str.contains('Styrene port inventory', na=False).idxmax()].reset_index(drop=True)
df_sm_weekly_f_inv = df_sm_weekly_f.iloc[df_sm_weekly_f.iloc[:, 0].str.contains('Styrene port inventory', na=False).idxmax():df_sm_weekly_f.iloc[:, 0].str.contains('Cash flow', na=False).idxmax()].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_sm_weekly_f_cf = df_sm_weekly_f.iloc[df_sm_weekly_f.iloc[:, 0].str.contains('Cash flow', na=False).idxmax():].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[1]).drop(d.index[:2]).reset_index(drop=True)).pipe(lambda d: d.rename(columns={d.columns[0]: 'Cash flow (yuan/mt)'}))

dfs = [df_bz_daily_f, df_bz_weekly_f_or, df_bz_weekly_f_inv, df_sm_weekly_f_or, df_sm_weekly_f_inv, df_sm_weekly_f_cf]

for df in dfs:
    cols = df.columns[1:]
    df[cols] = df[cols].replace(',', '', regex=True)
    df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')

print("✅ CCFGroup 테이블 DataFrame 추출 및 전처리 완료")

# =========================================================
# 8. EIA 데이터 및 환율 정보 스크래핑
# =========================================================
url_eia = "https://www.eia.gov/dnav/pet/hist_xls/WPULEUS3w.xls"

response = requests.get(url_eia)
response.raise_for_status()

df_eia = pd.read_excel(
    BytesIO(response.content),
    sheet_name="Data 1",
    skiprows=2,
    engine="calamine"
)

print(f"EIA 최근 데이터 - {df_eia.iloc[-1,-2]} : {df_eia.iloc[-1,-1]}")

USDCNY_str = html.fromstring(
    requests.get(
        "http://m.stock.naver.com/marketindex/exchangeWorld/USDCNY",
        headers={"User-Agent": "Mozilla/5.0"},
        verify=False
    ).text
).xpath(
    "/html/body/div[1]/div[1]/div[2]/div/div[1]/div[2]/div[2]/strong/text()"
)[0]

USDCNY = float(USDCNY_str.replace(',', ''))
usdcny_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
print(f"✅ 현재 USD/CNY 환율: {USDCNY} (추출 시간: {usdcny_time})")

# =========================================================
# 9. ICIS 데이터 로드 및 전처리
# =========================================================
today_dt = datetime.now()
start_date = (today_dt - timedelta(days=30)).strftime('%Y-%m-%d')
end_date = "2099-12-31"
dlkey = "cfdb4d7d-85c9-4334-9c0d-a7a3dea0e13f"

print(f"\n조회 기간: {start_date} ~ {end_date}")

urls_icis = {
    "ed": f"https://petchem.analytics.icis.com/wp-content/plugins/tschachsolutions/inc/highChartsInterface/api/hcapi.php?csv=1&dlkey={dlkey}&query_id=ICIS_petchem_margin_benzene_flexi_two_margin_comparison&form_hcg6980045b8bddb975559628%5B%5D={start_date}&form_hcg6980045b8bddb975559628%5B%5D={end_date}&form_hcg6980045b8bddb975559628%5B%5D=Weekly&form_hcg6980045b8bddb975559628%5B%5D=USD&form_hcg6980045b8bddb975559628%5B%5D=Tonne&form_hcg6980045b8bddb975559628%5B%5D=US%20Gulf&form_hcg6980045b8bddb975559628%5B%5D=Benzene%20extractive%20distillation&form_hcg6980045b8bddb975559628%5B%5D=Contract&form_hcg6980045b8bddb975559628%5B%5D=North%20East%20Asia&form_hcg6980045b8bddb975559628%5B%5D=Benzene%20extractive%20distillation&form_hcg6980045b8bddb975559628%5B%5D=Spot&chartset=comparison",
    "stdp": f"https://petchem.analytics.icis.com/wp-content/plugins/tschachsolutions/inc/highChartsInterface/api/hcapi.php?csv=1&dlkey={dlkey}&query_id=ICIS_petchem_margin_benzene_flexi_two_margin_comparison&div=hcg698d6e50f4221087364321&form_hcg698d6e50f4221087364321%5B%5D={start_date}&form_hcg698d6e50f4221087364321%5B%5D={end_date}&form_hcg698d6e50f4221087364321%5B%5D=Weekly&form_hcg698d6e50f4221087364321%5B%5D=USD&form_hcg698d6e50f4221087364321%5B%5D=Tonne&form_hcg698d6e50f4221087364321%5B%5D=US%20Gulf&form_hcg698d6e50f4221087364321%5B%5D=Benzene%20STDP&form_hcg698d6e50f4221087364321%5B%5D=Contract&form_hcg698d6e50f4221087364321%5B%5D=North%20East%20Asia&form_hcg698d6e50f4221087364321%5B%5D=Benzene%20STDP&form_hcg698d6e50f4221087364321%5B%5D=Spot&chartset=comparison",
    "hda": f"https://petchem.analytics.icis.com/wp-content/plugins/tschachsolutions/inc/highChartsInterface/api/hcapi.php?csv=1&dlkey={dlkey}&query_id=ICIS_petchem_margin_benzene_flexi_two_margin_comparison&div=hcg698d786e3e0ce133040510&form_hcg698d786e3e0ce133040510%5B%5D={start_date}&form_hcg698d786e3e0ce133040510%5B%5D={end_date}&form_hcg698d786e3e0ce133040510%5B%5D=Weekly&form_hcg698d786e3e0ce133040510%5B%5D=USD&form_hcg698d786e3e0ce133040510%5B%5D=Tonne&form_hcg698d786e3e0ce133040510%5B%5D=US%20Gulf&form_hcg698d786e3e0ce133040510%5B%5D=Benzene%20hydrodealkylation&form_hcg698d786e3e0ce133040510%5B%5D=Contract&form_hcg698d786e3e0ce133040510%5B%5D=North%20East%20Asia&form_hcg698d786e3e0ce133040510%5B%5D=Benzene%20hydrodealkylation&form_hcg698d786e3e0ce133040510%5B%5D=Spot&chartset=comparison"
}

results = {}

for key, url in urls_icis.items():
    try:
        df = pd.read_csv(url)
        df['Date'] = pd.to_datetime(df['Date'])
        results[key] = df.loc[df.groupby('Name')['Date'].idxmax()].reset_index(drop=True)
        print(f"✅ [{key.upper()}] 데이터 로드 및 전처리 완료")
    except pd.errors.EmptyDataError:
        print(f"❌ [{key.upper()}] 서버에서 빈 데이터를 반환했습니다.")
    except Exception as e:
        print(f"❌ [{key.upper()}] 예상치 못한 오류: {e}")

if results:
    df_icis = pd.concat(results.values(), ignore_index=True)
else:
    print("\n가져올 수 있는 데이터가 없어 병합을 수행하지 못했습니다.")

# =========================================================
# 10. 최종 결과 데이터프레임 매핑 (Value & Date)
# =========================================================
data_map_bz_1 = {
    "Asia BZ": df_bz_weekly_f_or.iloc[1, 2],
    "중국 BZ (Oil-based)": df_bz_weekly_f_or.iloc[0, 2],
    "중국 BZ (Coal-based)": df_bz_weekly_f_or.iloc[2, 2],
    "미국 Refinery": df_eia.iloc[-1,-1],
    "NEA BZ Extraction": df_icis.iloc[0,2],
    "NEA TDP": df_icis.iloc[2,2],
    "NEA Cracker": df_icis.iloc[4,2],
    "US BZ Extraction": df_icis.iloc[1,2],
    "US STDP": df_icis.iloc[3,2],
    "BZ RMB(M+1)-FOB KOR": 0,
    "BZ US(M+2)-FOB KOR": 0,
    "BZ CFR CHN-ARA(M)": 0,
    "SM ARA(M+1)-Asia": 0,
    "중국 D/S 복합 가동률": ((df_bz_weekly_f_or.iloc[3, 2]*0.38) + (df_bz_weekly_f_or.iloc[4, 2]*0.14) + (df_bz_weekly_f_or.iloc[6, 2] *0.24) + (df_bz_weekly_f_or.iloc[7, 2] *0.12))/0.88,
    "중국 SM 가동률": df_bz_weekly_f_or.iloc[3, 2],
    "중국 SM 마진": df_bz_daily_f.iloc[0, 3]/USDCNY,
    "중국 PS/EPS/ABS 복합 가동률": ((df_sm_weekly_f_or.iloc[2, 2]*0.29) + (df_sm_weekly_f_or.iloc[3, 2]*0.21) + (df_sm_weekly_f_or.iloc[1, 2] *0.24))/0.74,
    "중국 PS/EPS/ABS 복합 마진": 0,
    "중국 Phenol 가동률": df_bz_weekly_f_or.iloc[4, 2],
    "중국 Phenol 마진": df_bz_daily_f.iloc[1, 3]/USDCNY,
    "중국 Aniline 가동률": df_bz_weekly_f_or.iloc[7, 2],
    "중국 Aniline 마진": df_bz_daily_f.iloc[2, 3]/USDCNY,
    "중국 CPL 가동률": df_bz_weekly_f_or.iloc[6, 2],
    "중국 CPL 마진": df_bz_daily_f.iloc[3, 3]/USDCNY,
    "중국 BZ 화동": df_bz_weekly_f_inv.iloc[0, 2],
    "중국 SM 화동": df_sm_weekly_f_inv.iloc[0, 2]
}

df_bz_result = pd.Series(data_map_bz_1).to_frame(name='Value')

# ✅ 에러가 나던 분리 로직을 기존 원본 코드(df_bz_daily[3])로 정상 복구
data_map_bz_2 = {
    "Asia BZ": df_bz_weekly_f_or.columns[2],
    "중국 BZ (Oil-based)": df_bz_weekly_f_or.columns[2],
    "중국 BZ (Coal-based)": df_bz_weekly_f_or.columns[2],
    "미국 Refinery": df_eia.iloc[-1,-2].strftime('%Y-%m-%d'),
    "NEA BZ Extraction": df_icis.iloc[0,1].strftime('%Y-%m-%d'),
    "NEA TDP": df_icis.iloc[2,1].strftime('%Y-%m-%d'),
    "NEA Cracker": df_icis.iloc[4,1].strftime('%Y-%m-%d'),
    "US BZ Extraction": df_icis.iloc[1,1].strftime('%Y-%m-%d'),
    "US STDP": df_icis.iloc[3,1].strftime('%Y-%m-%d'),
    "BZ RMB(M+1)-FOB KOR": 0,
    "BZ US(M+2)-FOB KOR": 0,
    "BZ CFR CHN-ARA(M)": 0,
    "SM ARA(M+1)-Asia": 0,
    "중국 D/S 복합 가동률": df_bz_weekly_f_or.columns[2],
    "중국 SM 가동률": df_bz_weekly_f_or.columns[2],
    "중국 SM 마진": df_bz_daily[0].iloc[0, 0].split('(')[1].strip(')'),
    "중국 PS/EPS/ABS 복합 가동률": df_sm_weekly_f_or.columns[2],
    "중국 PS/EPS/ABS 복합 마진": 0,
    "중국 Phenol 가동률": df_bz_weekly_f_or.columns[2],
    "중국 Phenol 마진": df_bz_daily[0].iloc[0, 0].split('(')[1].strip(')'),
    "중국 Aniline 가동률": df_bz_weekly_f_or.columns[2],
    "중국 Aniline 마진": df_bz_daily[0].iloc[0, 0].split('(')[1].strip(')'),
    "중국 CPL 가동률": df_bz_weekly_f_or.columns[2],
    "중국 CPL 마진": df_bz_daily[0].iloc[0, 0].split('(')[1].strip(')'),
    "중국 BZ 화동": df_bz_weekly_f_inv.columns[2],
    "중국 SM 화동": df_sm_weekly_f_inv.columns[2]
}

df_bz_result['Date'] = df_bz_result.index.map(data_map_bz_2)
df_bz_result['Value'] = df_bz_result['Value'].astype(float).round(2)

# =========================================================
# 11. 엑셀 파일로 출력 (xlsxwriter)
# =========================================================
today_str = datetime.now().strftime('%Y-%m-%d')
file_name = f"bz_result_{today_str}.xlsx"

for col in df_eia.select_dtypes(include=['datetime64', 'datetimetz']).columns:
    df_eia[col] = df_eia[col].dt.strftime('%Y-%m-%d')

if 'df_icis' in locals():
    for col in df_icis.select_dtypes(include=['datetime64', 'datetimetz']).columns:
        df_icis[col] = df_icis[col].dt.strftime('%Y-%m-%d')

df_usdcny = pd.DataFrame({
    'USDCNY': [USDCNY],
    'Extraction Time': [usdcny_time]
})

targets = {
    "bz_result": df_bz_result,
    "eia": df_eia,
    "icis": df_icis if 'df_icis' in locals() else pd.DataFrame(),
    "USDCNY": df_usdcny,
    "url": df_url
}

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    workbook = writer.book
    
    border_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#BCD1E4', 'border': 1, 'align': 'center'})

    for sheet_name, df in targets.items():
        if df is not None and not df.empty:
            df.to_excel(writer, sheet_name=sheet_name, index=True)
            worksheet = writer.sheets[sheet_name]
            
            rows, cols = df.shape
            worksheet.conditional_format(0, 0, rows, cols, {
                'type': 'formula',
                'criteria': '=1',
                'format': border_format
            })
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num + 1, str(value), header_format)
            worksheet.write(0, 0, str(df.index.name or 'Index'), header_format)

            idx_max_len = max(df.index.astype(str).map(len).max(), len(str(df.index.name or "Index"))) + 3
            worksheet.set_column(0, 0, idx_max_len)
            
            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 3
                worksheet.set_column(i + 1, i + 1, column_len)

print(f"✅ '{file_name}' 저장 완료!")

# =========================================================
# 12. 이메일 발송 (Gmail 앱 비밀번호 사용)
# =========================================================
print("=== 메일 발송 준비 ===")

sender_email = os.environ.get("GMAIL_USER")
app_password = os.environ.get("GMAIL_APP_PASSWORD")

to_emails = "carly1206@sk.com, rchangjo@sk.com"
cc_emails = "michael.park@sk.com, jsoh@sk.com, hoseok@sk.com, hyo548@sk.com, cr7@sk.com, jp_lee@sk.com"
# to_emails = "jp_lee@sk.com"
# cc_emails = "jp_lee@sk.com"

subject = f"BZ CCF {today_str}"

html_table = df_bz_result.to_html(justify='center', index=True)

custom_table_tag = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; text-align:center; font-family:Calibri, Arial, sans-serif; font-size:13px;">'
html_table = html_table.replace('<table border="1" class="dataframe">', custom_table_tag)

html_body = f"""
<html>
<body style="margin:0; padding:0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
            <td style="padding:20px;
                       font-family:Calibri, Arial, sans-serif;
                       font-size:14px;
                       color:#000000;">

                안녕하세요,<br><br>

                오늘자 BZ CCF 추출 결과입니다.<br>
                상세 내용은 첨부파일을 확인해 주시기 바랍니다.<br><br>

                {html_table}

            </td>
        </tr>
    </table>
</body>
</html>
"""

msg = EmailMessage()
msg['Subject'] = subject
msg['From'] = sender_email
msg['To'] = to_emails
msg['Cc'] = cc_emails
msg.set_content("HTML 뷰어를 지원하는 메일 클라이언트를 사용해 주세요.") 
msg.add_alternative(html_body, subtype='html')

with open(file_name, 'rb') as f:
    excel_data = f.read()
    
msg.add_attachment(
    excel_data, 
    maintype='application', 
    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
    filename=file_name
)

if sender_email and app_password:
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)
        print("✅ 이메일 발송 완료!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")
else:
    print("⚠️ GMAIL_USER 또는 GMAIL_APP_PASSWORD 환경변수가 설정되지 않아 메일을 발송하지 않았습니다.")

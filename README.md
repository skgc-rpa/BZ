# 📈 BZ & SM CCF Market Data Auto-Scraper

> **중국 BZ/SM 시장(CCFGroup) 데이터 수집 및 글로벌 지표(EIA/ICIS) 통합 분석 자동화 파이프라인**

본 프로젝트는 **CCFGroup**에서 Benzene(BZ) 및 Styrene Monomer(SM) 관련 시장 데이터를 자동으로 수집하고, **EIA(미국 에너지정보청)**의 정제소 데이터와 **ICIS**의 글로벌 마진 지표를 결합하여 비즈니스 리포트를 생성하는 **Python 기반 자동화 시스템**입니다.

---

## ✨ 주요 기능 (Key Features)

* **🌐 다각적 데이터 통합**: CCFGroup의 중국 내수 가격뿐만 아니라 EIA의 미국 가동률, ICIS의 지역별 마진 데이터를 통합하여 수집합니다.
* **📊 지능형 데이터 전처리**: Pandas를 활용해 Raw 데이터를 정제하며, 실시간 환율(USD/CNY)을 반영하여 글로벌 단위의 마진을 자동 산출합니다.
* **🎨 맞춤형 엑셀 리포트**: `xlsxwriter`를 통해 기업용 포맷(조건부 서식, 헤더 강조, 셀 너비 최적화)이 적용된 다중 시트 보고서를 자동 생성합니다.
* **📧 스마트 이메일 알림**: 생성된 분석 결과 요약 표를 메일 본문에 포함하고, 상세 엑셀 파일을 첨부하여 지정된 수신처로 자동 발송합니다.
* **🤖 Cloud 기반 무중단 운영**: **GitHub Actions**를 통해 매주 평일(월~금) 오전 10시(KST)에 별도의 서버 없이 클라우드에서 자동 실행됩니다.

---

## 📁 저장소 구조 (Repository Structure)

```text
BZ/
├── .github/
│   └── workflows/
│       └── daily_ccf.yml         # CI/CD 스케줄러 (GitHub Actions)
├── main.py                       # 핵심 로직 (스크래핑, 연산, 메일 발송)
├── requirements.txt              # 의존성 패키지 목록
└── README.md                     # 프로젝트 가이드
```

---

## ⚙️ 환경 설정 (Setup & Installation)

### 1. GitHub Secrets 등록 (필수)
자동 이메일 발송 기능을 위해 GitHub 저장소의 `Settings > Secrets and variables > Actions` 메뉴에 아래 두 가지 변수를 반드시 등록해야 합니다.

| 변수명 | 설명 | 비고 |
| :--- | :--- | :--- |
| **`GMAIL_USER`** | 발송용 Gmail 주소 | 예: `example@gmail.com` |
| **`GMAIL_APP_PASSWORD`** | Gmail 앱 비밀번호 | 16자리 (띄어쓰기 없이 입력) |

### 2. 로컬 개발 환경 구축
로컬 PC에서 스크립트를 수동으로 테스트하거나 수정할 경우 아래 라이브러리들을 설치해야 합니다.

```bash
# 의존성 패키지 설치
pip install requests pandas beautifulsoup4 lxml xlsxwriter urllib3
```
> **Note:** 실행 시 환경 변수가 로컬에 설정되어 있지 않으면 메일 발송은 건너뛰고 엑셀 파일만 생성됩니다.

---

## ⏰ 실행 스케줄 (Schedule)

GitHub Actions를 통해 설정된 자동 실행 주기입니다.

* **실행 주기**: 매주 월요일 ~ 금요일 (평일)
* **실행 시간**: KST 오전 10:00 (UTC 01:00)
* **구동 방식**: GitHub Actions Ubuntu 서버 워크플로우를 통한 자동 실행
* **Cron 설정**: `0 1 * * 1-5`

---

## 🛠 기술 스택 (Tech Stack)

* **Language**: Python 3.x
* **Data Analysis**: Pandas, NumPy
* **Scraping**: Requests, BeautifulSoup4, Lxml
* **Reporting**: XlsxWriter, SMTPLib
* **Automation**: GitHub Actions

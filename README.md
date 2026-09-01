# GMP Automation System (GMP 환경측정 자동화 시스템)

제약 및 바이오 시설의 환경측정 PDF 기록서를 분석하여 평균값, 기준선(경고/조치), 차트가 포함된 Microsoft Excel 보고서로 자동 변환하는 로컬 웹 애플리케이션입니다.

## 주요 기능

- **다중 PDF 일괄 처리:** 여러 AHU의 PDF 기록서를 한 번에 업로드하여 처리합니다.
- **유연한 OCR 방식 지원:** API 기반 Online OCR과 셀프 호스팅 OCR 서버 기반 Offline OCR을 모두 지원합니다.
- **자동 Excel 보고서 생성:** 시험 항목별로 데이터 시트, 요약 테이블, 피벗/차트 시트를 자동으로 생성합니다.
- **기준 초과 시각화:** 설정된 경고/조치 기준을 벗어난 측정값을 빨간색 하이라이트로 표시합니다.
- **AHU 및 반기별 자동 정렬:** 공조기(AHU) 번호와 측정 반기(상/하반기) 기준 데이터 자동 그룹화.

## 지원되는 측정 항목

| 코드 | 측정 항목 | 생성 Excel 파일명 |
| --- | --- | --- |
| A | 부유입자 측정 (Airborne Particle Test) | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | 풍속 측정 (Air Velocity Test) | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | 환기횟수 측정 (Air Change Rate Test) | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA 필터 검사 (HEPA Filter Test) | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | 기류패턴 시험 (Airflow Pattern Test) | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## 사전 요구 사항

- Python 3.10 이상
- Poppler (PDF 페이지를 이미지로 변환하기 위해 필요)
- OCR 실행을 위한 네트워크 연결
- 다음 OCR 모드 중 하나 선택:
  - Anthropic API Key를 사용하는 **Online OCR**
  - 셀프 호스팅 OCR 엔드포인트 URL을 사용하는 **Offline OCR**

### Poppler 설치 방법:

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

Windows의 경우 [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)를 다운로드하여 `C:\poppler`에 압축 해제 후, `C:\poppler\Library\bin` 경로를 System `PATH` 환경 변수에 추가하세요.

## 설치 방법

```bash
git clone <REPOSITORY_URL>
cd gmp_automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows 환경에서는 아래 명령어로 가상환경을 활성화합니다:

```bat
.venv\Scripts\activate
```

## 애플리케이션 실행

```bash
python app.py
```

또는 제공된 실행 스크립트를 사용할 수 있습니다:

```bash
# macOS/Linux
bash START_LINUX.sh

# Windows
START_WINDOWS.bat
```

웹 브라우저에서 `http://localhost:5001` 접속 시 자동으로 Online OCR 페이지(`http://localhost:5001/online`)로 이동합니다.

## Docker 배포

Docker Engine 및 Docker Compose 플러그인이 필요합니다. 아래 명령어로 실행할 수 있습니다:

```bash
bash deploy.sh
```

기본 웹 애플리케이션은 `http://localhost:5000/`에서 접속 가능합니다. 포트를 변경해야 하는 경우:

```bash
HOST_PORT=8080 bash deploy.sh
```

생성된 Excel 파일과 임시 업로드 파일은 Docker 볼륨에 저장됩니다. 로그 확인 및 정지는 아래 명령어를 사용하세요:

```bash
docker compose -p gmp-main logs -f
docker compose -p gmp-main down
```

## 사용 방법

1. `/online` (Online OCR) 또는 `/offline` (Offline OCR) 페이지로 이동합니다. 상단 상호 이동 버튼으로 쉽게 전환할 수 있습니다.
2. 선택한 모드에 따라 Anthropic API Key 또는 Offline OCR 엔드포인트 URL을 입력합니다.
3. 측정할 시험 항목(A~E) 중 하나를 선택합니다.
4. 동일한 시험 항목의 PDF 파일(들)을 업로드합니다.
5. `Excel 자동 생성 시작` 버튼을 클릭하여 보고서를 다운로드합니다.

*각 PDF 파일은 1개 AHU의 1개 반기 측정 기록서여야 합니다. OCR 인식률은 스캔된 PDF의 화질에 영향을 받습니다.*

## Kaggle 기반 Offline OCR 사용 안내

Offline OCR은 API 비용 없이 사용할 수 있는 대안입니다. GPU 및 인터넷이 활성화된 Kaggle 환경에서 서버를 실행하세요: [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server).

1. [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server) 노트를 엽니다.
2. Accelerator를 GPU로 설정하고 Internet 옵션을 Turn on 합니다.
3. Kaggle Secrets에 `NGROK_AUTHTOKEN` 이름으로 본인의 ngrok 토큰을 추가합니다.
4. 모든 셀을 실행합니다.
5. 화면에 출력된 `https://*.ngrok-free.app` URL을 복사하여 웹 UI의 Offline OCR 엔드포인트에 입력합니다.

*Kaggle 세션과 ngrok URL은 임시적이므로 재실행 시 URL이 변경됩니다. 데이터 추출 테스트 및 검증 시에는 `debug_ocr.py`를 활용할 수 있습니다:*

```bash
python debug_ocr.py <PDF_파일_경로> <NGROK_URL>
```

## 프로젝트 구조

```text
gmp_automation/
├── app.py                 # Flask 웹 애플리케이션 및 엔드포인트
├── config.py              # 시험별 기준값 및 환경 설정
├── ocr_engine.py          # Online OCR 기반 PDF 데이터 추출 엔진
├── deepseek_ocr/          # Offline OCR 클라이언트, 파서 및 서버 코드
├── excel_generator.py     # Excel 보고서 및 차트 자동 생성 모듈
├── templates/index.html   # 웹 인터페이스 템플릿
├── boilerplate/           # Excel 양식 템플릿
├── uploads/               # PDF 업로드 임시 저장소
├── outputs/               # 생성된 Excel 보고서 저장소
├── debug_ocr.py           # Offline OCR 디버깅 도구
└── requirements.txt       # Python 의존성 패키지 목록
```

## 참고 사항

- Online OCR 사용 시 Anthropic 계정에 따라 API 비용이 발생합니다.
- 각 시험 항목별 기준값 및 설정은 `config.py`에서 변경할 수 있습니다.
- 1회 업로드 용량 제한은 100MB입니다.
- Excel 보고서가 성공적으로 생성되면 업로드된 임시 PDF 파일은 자동으로 삭제됩니다.

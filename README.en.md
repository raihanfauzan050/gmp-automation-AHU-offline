[![Korean README](https://img.shields.io/badge/README-한국어-2563eb)](README.md)

# GMP Automation System - Offline

A local web application that converts pharmaceutical-facility environmental measurement PDFs into structured Microsoft Excel reports with averages, limit indicators, and charts.

## Features

- Processes multiple PDFs for different AHUs in one request.
- Extracts PDF data through a self-hosted DeepSeek OCR endpoint.
- Produces an Excel workbook for each test type with data sheets, summary tables, and charts.
- Highlights values outside configured limits in red.
- Groups results by AHU and measurement semester.

## Supported Tests

| Code | Test | Excel File |
| --- | --- | --- |
| A | Airborne Particle Test | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | Air Velocity Test | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | Air Change Rate Test | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA Filter Test | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | Airflow Pattern Test | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## Requirements

- Python 3.10 or later
- Poppler, used to convert PDF pages into images
- Docker Desktop is optional and only required for container deployment
- At least 16 GB of system RAM. The local model uses CPU inference for GTX 1050 compatibility and may be slow.

### Installing Poppler

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

On Windows, download [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases), extract it to `C:\poppler`, and add `C:\poppler\Library\bin` to the system `PATH` environment variable.

## Installation

```bash
git clone --branch gmp-offline --single-branch https://github.com/irfanqs/gmp-automation.git
cd gmp-automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, activate the virtual environment with:

```bat
.venv\Scripts\activate
```

## Check Windows GPU Compatibility

Before setting up local DeepSeek OCR, run `CHECK_HARDWARE.bat`. It detects display adapters, NVIDIA driver availability, and VRAM without sending hardware data anywhere.

The current deployment uses CPU inference, so it works without a supported GPU. The script reports NVIDIA driver and VRAM information for future CUDA acceleration; GTX 1050 and integrated GPUs continue to use CPU mode.

## Run Without Docker on Windows

Run `START_WINDOWS.bat`. It creates `.venv`, installs the CPU OCR dependencies, starts the local DeepSeek OCR model, waits for it to finish loading, and then starts the GMP web application at `http://localhost:5002/offline`.

The model starts in a separate `GMP Offline OCR Model` window. Keep that window open while using the application. The first startup downloads the model and can take a long time on CPU; later startups reuse the downloaded model cache.

## Docker Deployment

Docker Desktop on Windows or Docker Engine with the Docker Compose plugin is required. Deploy this branch with:

```bash
bash deploy.sh
```

The first deployment downloads DeepSeek OCR from Hugging Face and loads it into CPU memory. This can take a long time and requires an internet connection only for the initial download. Docker stores the model in a persistent volume, so later starts do not download it again.

When the `ocr` service becomes healthy, the Offline OCR application is available at `http://localhost:5002/offline`. No URL or API key is entered in the web interface. Follow the initial model download and loading progress with:

```bash
docker compose -p gmp-offline logs -f ocr
```

To use a different host port:

```bash
HOST_PORT=8082 bash deploy.sh
```

Generated files and temporary uploads are stored in Docker volumes. Use the following commands to inspect logs or stop the application:

```bash
docker compose -p gmp-offline logs -f
docker compose -p gmp-offline down
```

## Usage

1. Wait until the local `ocr` service has finished loading the model.
2. Select one of the supported test types (A-E).
3. Upload one or more PDFs of the same test type.
4. Click `Start Excel Generation` to generate and download the report.

Each PDF must contain the measurement record for one AHU and one semester. OCR accuracy depends on the quality of the scanned PDF.

## Local CPU OCR

The current setup intentionally uses CPU inference so it can run without a compatible CUDA GPU. GTX 1050, Intel/AMD integrated graphics, and other unsupported GPUs are not used for inference. CPU OCR is expected to be slow, especially for multi-page PDFs.

The first model download requires internet access. After it is cached in Docker, the OCR service does not use Kaggle, ngrok, or an external OCR API. Use `debug_ocr.py` to inspect raw OCR text when extraction needs to be reviewed:

```bash
python debug_ocr.py <PATH_TO_PDF>
```

## Project Structure

```text
gmp-automation/
├── app.py                 # Flask application and endpoints
├── config.py              # Test limits and environment configuration
├── deepseek_ocr/          # Local OCR client, parser, and model server
├── excel_generator.py     # Excel report and chart generator
├── templates/index.html   # Web interface template
├── boilerplate/           # Example Excel templates
├── uploads/               # Temporary PDF uploads
├── outputs/               # Generated Excel reports
├── debug_ocr.py           # Offline OCR debugging tool
└── requirements.txt       # Python dependencies
```

## Notes

- The first model download may take time and consume significant disk space.
- Test limits and settings can be changed in `config.py`.
- The upload limit is 100 MB per request.
- Temporary uploaded PDFs are deleted automatically after an Excel report is generated successfully.

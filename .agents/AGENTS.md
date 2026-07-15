# ⚓ Quy Tắc Dự Án - OptimoorTool

Tài liệu này định nghĩa các quy tắc lập trình, tiêu chuẩn thiết kế và hành vi Agent khi làm việc trong dự án **OptimoorTool** — công cụ xử lý báo cáo phân tích mooring từ phần mềm Optimoor.

---

## 🛠️ Công Nghệ & Thư Viện Sử Dụng

- **Ngôn ngữ chính**: Python 3.9+
- **Framework giao diện**: Streamlit `>= 1.58.0`
- **Thư viện phân tích dữ liệu**:
  - `pandas` — xử lý bảng dữ liệu mooring
  - `numpy` — tính toán số học
  - `matplotlib` + `seaborn` — vẽ biểu đồ Rose Chart
  - `openpyxl` — đọc/ghi file Excel (`.xlsx`, `.xlsm`)
  - `scipy` — tính toán thống kê (nếu cần)
  - `pillow` — xử lý hình ảnh
  - `python-docx` — tạo báo cáo Word (.docx)
- **Module nội bộ**:
  - `optimoor_rtf_to_excel.py` — parser RTF → Excel
  - `optimoor_report_generator.py` — sinh báo cáo Word từ input .xlsm + output .xlsx
  - `app.py` — entry point giao diện Streamlit
- **Quy tắc thư viện**:
  - Không tự ý cài thêm thư viện nếu không có sự đồng ý của người dùng.
  - Luôn cập nhật `requirements.txt` khi bổ sung dependency mới.

---

## 📂 Cấu Trúc Thư Mục

```
OptimoorTool/
├── app.py                          # Entry point Streamlit — giao diện chính
├── optimoor_rtf_to_excel.py        # Module parser RTF → Excel (Optimoor reports)
├── optimoor_report_generator.py    # Module sinh báo cáo Word từ 2 file Excel
├── requirements.txt                # Danh sách thư viện
├── scripts/
│   └── create_report_template.py   # Tạo template/Mooring_Report_Template.docx
├── template/
│   ├── Optimoor Tool-Master_Post-Processing.xlsm  # Template Excel gốc
│   └── Mooring_Report_Template.docx               # Template báo cáo Word chuẩn
├── tasks/
│   ├── mooring_force/
│   │   ├── input/                  # File đầu vào Rose Chart (Excel)
│   │   └── output/                 # Biểu đồ đầu ra (PNG/PDF)
│   ├── optimoor_rtf_to_excel/
│   │   ├── input/                  # File .RTF tải lên
│   │   └── output/                 # File .xlsx kết quả
│   └── generate_report/
│       ├── input/                  # OptimoorTool_Input.xlsm + OutputResult.xlsx
│       └── output/                 # File .docx báo cáo
├── output_charts/                  # Thư mục lưu chart tạm thời
├── images/
│   └── template_input_data.xlsx    # File mẫu dữ liệu đầu vào
└── OPTIMOOR-TECHNICAL WORKFLOW GUIDE.html  # Tài liệu hướng dẫn kỹ thuật
```

> Không thay đổi tên thư mục `tasks/`, `template/`, `images/` vì chúng được hard-code trong `app.py` thông qua các hằng số `Path`.

---

## 🔧 Hằng Số Đường Dẫn (app.py)

Các đường dẫn được định nghĩa ở đầu `app.py` và phải được giữ nguyên:

```python
BASE_DIR            = Path(__file__).resolve().parent
OUTPUT_DIR          = BASE_DIR / "output_charts"
TASKS_DIR           = BASE_DIR / "tasks"
MOORING_TASK_DIR    = TASKS_DIR / "mooring_force"
MOORING_INPUT_DIR   = MOORING_TASK_DIR / "input"
MOORING_OUTPUT_DIR  = MOORING_TASK_DIR / "output"
OPTI_TASK_DIR       = TASKS_DIR / "optimoor_rtf_to_excel"
OPTI_INPUT_DIR      = OPTI_TASK_DIR / "input"
OPTI_OUTPUT_DIR     = OPTI_TASK_DIR / "output"
REPORT_TASK_DIR     = TASKS_DIR / "generate_report"
REPORT_INPUT_DIR    = REPORT_TASK_DIR / "input"
REPORT_OUTPUT_DIR   = REPORT_TASK_DIR / "output"
TEMPLATE_DIR        = BASE_DIR / "template"
REPORT_TEMPLATE_PATH = TEMPLATE_DIR / "Mooring_Report_Template.docx"
```

Khi thêm task mới, hãy theo cùng pattern đặt tên và tổ chức `input/output`.

---

## 📄 Module `optimoor_rtf_to_excel.py`

- **Mục đích**: Parse file `.RTF` xuất từ phần mềm Optimoor thành file `.xlsx`.
- **Quy tắc quan trọng**:
  - RTF encoding sử dụng **cp1252** (Windows-1252) — không thay đổi.
  - `NUMBER_PATTERN = r"[-+]?\d*\.\d+|[-+]?\d+"` — dùng regex này để trích xuất số từ report.
  - Hàm `rtf_to_text()` và `_decode_rtf_line()` là core parser — chỉ chỉnh sửa nếu thực sự cần thiết và phải test kỹ.
  - File RTF được split thành các **Batch Run** (pattern `Batch Run N:`).
  - Luôn kiểm tra cả `.rtf` và `.RTF` (case-insensitive) khi glob file.

---

## 📄 Module `optimoor_report_generator.py`

- **Mục đích**: Sinh báo cáo Word (`.docx`) từ cặp file **OptimoorTool_Input.xlsm** + **OutputResult.xlsx**.
- **Input sheets** (`.xlsm`): `1A.Berth_Data`, `2A.Vessel_Data`, `3A.GeneralEnvir.`, `3B.RunCase`.
- **Output sheets** (`.xlsx`): `Cases`, `Greatest_Line`, `Greatest_Bollard`, `Hook_Bollard_Forces`, `Line_Tensions`.
- **Cấu trúc báo cáo**: Cover → Input Data → Design Scenarios → Analysis Results → Summary KPI.
- **API chính**: `generate_report()`, `validate_input_files()`, `build_report_context()`.
- **Ngưỡng OCIMF**: `MBL_LIMIT = 0.50` (50% MBL) dùng cho status OK/NG.
- **Template**: `template/Mooring_Report_Template.docx` — tái tạo bằng `python scripts/create_report_template.py`.

---

## 🎨 Tiêu Chuẩn Giao Diện Streamlit

- **Màu sắc biểu đồ** chuẩn (giữ nguyên thứ tự):
  ```python
  BASE_COLORS  = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6',
                  '#1abc9c', '#e67e22', '#34495e', '#e91e63', '#00bcd4']
  BASE_MARKERS = ['o', 's', '^', 'D', 'v', 'P', '*', 'X', 'h', '8']
  ```
- **DPI** mặc định cho chart: `DPI = 150`.
- **SHIP_ZOOM**: `0.10` (zoom level ảnh tàu trên chart).
- **DANGER_ZONE**: `(225, 315)` độ — vùng nguy hiểm trên Rose Chart.
- Các hằng số trên không được thay đổi trừ khi có yêu cầu rõ ràng từ người dùng.

---

## 🗂️ Quy Tắc Xử Lý File

- **Upload**: Luôn dùng `save_uploaded_files()` để lưu file tải lên vào đúng thư mục `input/`.
- **Cleanup**: Xóa file cũ trong `input/` và `output/` trước mỗi lần chạy mới (đã có trong `convert_rtf_to_excel()`).
- **Cross-platform**: Hàm `open_path()` và `can_open_file_locally()` xử lý mở file cho Windows/macOS/Linux — không bypass logic này.
- **File naming**: Dùng `sanitize_filename()` khi tạo tên file output để tránh ký tự đặc biệt.

---

## 💾 Quản Lý State Streamlit

- Dùng `st.session_state` để lưu trạng thái giữa các lần rerun.
- Gọi `ensure_task_directories()` ở đầu mỗi tác vụ để đảm bảo thư mục tồn tại.
- Không tạo thư mục trực tiếp ngoài hàm `ensure_task_directories()` hoặc `save_uploaded_files()`.

---

## 📝 Quy Tắc Code

- **Type hints**: Ưu tiên sử dụng type hints cho tất cả các hàm mới.
- **Docstring**: Viết docstring tiếng Anh cho hàm mới, comment giải thích logic phức tạp có thể bằng tiếng Việt.
- **Error handling**: Dùng `st.error()`, `st.warning()`, `st.success()` cho feedback giao diện — không dùng `print()`.
- **Excel output**: Luôn dùng `openpyxl` cho `.xlsx`; template `.xlsm` (macro-enabled) chỉ đọc bằng `openpyxl` với `keep_vba=True`.

---

## 🚀 Chạy & Phát Triển

```bash
# Cài thư viện
pip install -r requirements.txt

# Chạy ứng dụng
streamlit run app.py
```

- Môi trường ảo nằm ở `.venv/` — ưu tiên dùng môi trường này.
- Không commit thư mục `.venv/`, `__pycache__/`, `tasks/*/input/`, `tasks/*/output/`, `output_charts/` vào Git.

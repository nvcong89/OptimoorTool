"""Generate template/Mooring_Report_Template.docx sample from example files."""

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import optimoor_report_generator as gen

TEMPLATE_PATH = BASE_DIR / "template" / "Mooring_Report_Template.docx"
EXAMPLE_INPUT = BASE_DIR / "example" / "OptimoorTool_Input.xlsm"
EXAMPLE_OUTPUT = BASE_DIR / "example" / "OutputResult.xlsx"


def create_template() -> Path:
    """Write a sample report template populated from example Excel files."""
    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not EXAMPLE_INPUT.exists() or not EXAMPLE_OUTPUT.exists():
        context = gen.build_report_context(
            gen.InputReportData(),
            gen.OutputReportData(),
            gen.ReportMetadata(reference="TEMPLATE"),
        )
    else:
        input_data = gen.read_input_workbook(EXAMPLE_INPUT)
        output_data = gen.read_output_workbook(EXAMPLE_OUTPUT)
        context = gen.build_report_context(
            input_data,
            output_data,
            gen.ReportMetadata(
                reference="SAMPLE",
                project_name="Gemalink Container Terminal Phase 2 (Sample)",
            ),
        )

    doc = gen._render_document(context)
    doc.save(str(TEMPLATE_PATH))
    return TEMPLATE_PATH


if __name__ == "__main__":
    path = create_template()
    print(f"Created: {path}")

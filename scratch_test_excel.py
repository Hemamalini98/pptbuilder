import traceback
import os
from excel_report import create_excel_report

# Setup paths based on server.py
UPLOAD_DIR = os.path.abspath("uploads")
input_path = os.path.join(UPLOAD_DIR, "uploaded_content.pptx")
output_path = os.path.join(UPLOAD_DIR, "styled_output.pptx")
excel_path = os.path.join(UPLOAD_DIR, "test_report.xlsx")
extracts_dir = os.path.join(UPLOAD_DIR, "pdf_extracts")

if not os.path.exists(input_path):
    # Try finding any pptx in uploads
    for f in os.listdir(UPLOAD_DIR):
        if f.endswith(".pptx") and f != "styled_output.pptx":
            input_path = os.path.join(UPLOAD_DIR, f)
            break

print("Input path:", input_path)
print("Output path:", output_path)

try:
    create_excel_report(
        input_pptx=input_path,
        output_pptx=output_path,
        extracts_dir=extracts_dir,
        customer_name="Test Customer",
        project_name="Test Project",
        output_excel_path=excel_path
    )
    print("Success!")
except Exception as e:
    traceback.print_exc()

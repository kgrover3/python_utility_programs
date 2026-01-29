import os


PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))


XML_INPUT_DIR = os.path.join(PROJECT_ROOT, "xml_input")
XLSX_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "xlsx_output")


print(PROJECT_ROOT)
import os
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from config import XML_INPUT_DIR, XLSX_OUTPUT_DIR, PROJECT_ROOT


os.makedirs(XLSX_OUTPUT_DIR, exist_ok=True)

# Where to log oversized/skipped files
LOG_FILE = os.path.join(PROJECT_ROOT, "skipped_large_files.txt")

def log_skipped_file(filename, reason):
    """Append to log file with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {filename} - {reason}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(log_line)
    print(f"  SKIPPED & LOGGED: {filename} → {reason}")


def count_upc_elements(filepath, max_upc_estimate=1200000):
    """
    Quick low-memory estimate of <upc> count using iterparse.
    Returns count or None if parsing fails.
    """
    try:
        count = 0
        context = ET.iterparse(filepath, events=("end",))
        for event, elem in context:
            if elem.tag == "upc":
                count += 1
                if count > max_upc_estimate:
                    return count  # early exit
            elem.clear()  # free memory
        return count
    except Exception:
        return None


def parse_contact_lens_xml(filepath):
    """
    Parse both old and new catalog XML formats.
    One row per <upc>, with all known fields (old + new).
    Returns DataFrame or None if failed/skipped.
    """
    filename = os.path.basename(filepath)

    # Quick size pre-check (fast fail for huge files)
    upc_count = count_upc_elements(filepath)
    if upc_count is not None and upc_count > 1000000:
        log_skipped_file(filename, f"Too many UPCs (~{upc_count:,} > 1M limit)")
        return None

    try:
        tree = ET.parse(filepath)
        root = tree.getroot()

        rows = []

        for manuf in root.findall('.//manufacturer'):
            mcode = manuf.findtext('mCode', default='')
            mdesc = manuf.findtext('mDesc', default='')

            for prod in manuf.findall('product'):
                # Common fields (both formats)
                pcode     = prod.findtext('pCode', default='')
                pdesc     = prod.findtext('pDesc', default='')
                mode      = prod.get('mode', default='')
                qty       = prod.findtext('qty', default='')
                qtyunit   = prod.findtext('qtyUnit', default='')

                # Newer-format fields (safe: empty if missing)
                ptrialrev = prod.findtext('pTrialOrRev', default='')
                pmodality = prod.findtext('pModality', default='')
                ptype     = prod.findtext('pType', default='')

                for upc in prod.findall('upc'):
                    row = {
                        'Manufacturer_Code': mcode,
                        'Manufacturer_Desc': mdesc,
                        'Product_Code':      pcode,
                        'Product_Desc':      pdesc,
                        'Product_Mode':      mode,
                        'Trial_or_Revenue':  ptrialrev,      # T = trial, etc.
                        'Modality':          pmodality,      # WEEKLY, etc.
                        'Product_Type':      ptype,
                        'Quantity':          qty,
                        'Quantity_Unit':     qtyunit,

                        # UPC attributes
                        'UPC_ID':     upc.get('id', ''),
                        'Power':      upc.get('power', ''),
                        'Base_Curve': upc.get('basecurve', ''),
                        'Diameter':   upc.get('diameter', ''),
                        'Color':      upc.get('color', ''),
                        'Color2':     upc.get('color2', ''),
                        'Cylinder':   upc.get('cylinder', ''),
                        'Axis':       upc.get('axis', ''),
                        'Design':     upc.get('design', ''),
                        'Add':        upc.get('add', '')
                    }
                    rows.append(row)

        if not rows:
            print(f"  No <upc> elements found in {filename}")
            return None

        df = pd.DataFrame(rows)

        # Consistent column order for old + new files
        column_order = [
            'Manufacturer_Code', 'Manufacturer_Desc',
            'Product_Code', 'Product_Desc', 'Product_Mode',
            'Trial_or_Revenue', 'Modality', 'Product_Type',
            'Quantity', 'Quantity_Unit',
            'UPC_ID', 'Power', 'Base_Curve', 'Diameter',
            'Color', 'Color2', 'Cylinder', 'Axis', 'Design', 'Add'
        ]

        # Only include columns that actually exist
        available_cols = [col for col in column_order if col in df.columns]
        df = df[available_cols]

        # Final row-limit check (Excel max)
        if len(df) > 1048576:
            log_skipped_file(filename, f"Too many rows ({len(df):,} > Excel max 1,048,576)")
            return None

        return df

    except MemoryError:
        log_skipped_file(filename, "MemoryError - file too large for ElementTree/pandas")
        return None
    except ET.ParseError as e:
        log_skipped_file(filename, f"XML parse error: {e}")
        return None
    except Exception as e:
        log_skipped_file(filename, f"Unexpected error: {type(e).__name__} - {str(e)}")
        return None


# ────────────────────────────────────────────────
# Main loop
# ────────────────────────────────────────────────
print(f"Processing XML files from: {XML_INPUT_DIR}")
print(f"Output to: {XLSX_OUTPUT_DIR}")
print(f"Large/skipped files logged to: {LOG_FILE}\n")

for filename in os.listdir(XML_INPUT_DIR):
    if not filename.lower().endswith(('.xml', '.XML')):
        continue

    input_path = os.path.join(XML_INPUT_DIR, filename)
    print(f"Processing: {filename}")

    df = parse_contact_lens_xml(input_path)

    if df is not None and not df.empty:
        base_name = os.path.splitext(filename)[0]
        output_filename = f"{base_name}.xlsx"
        output_path = os.path.join(XLSX_OUTPUT_DIR, output_filename)

        try:
            df.to_excel(output_path, index=False, engine="openpyxl")
            print(f"  Saved → {output_filename}  ({len(df)} rows, {len(df.columns)} columns)")
        except Exception as e:
            log_skipped_file(filename, f"Excel write failed: {type(e).__name__} - {str(e)}")
    else:
        print("  Skipped")

print("\nProcessing complete.")
if os.path.exists(LOG_FILE):
    print(f"Check {LOG_FILE} for any skipped/large files.")
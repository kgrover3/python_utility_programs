# src/utils/writers.py
import os
# import pandas as pd

def write_xlsx_file(
    *,
    data: dict,
    filename: str,
    archive_path: str,
) -> str:
    """
    Writes data to xlsx file
    """
    try:
        os.makedirs(archive_path, exist_ok=True)
        file_path = os.path.join(archive_path, filename)
        
        
            
        return file_path
    except Exception as exc:
        raise RuntimeError(
            f"Failed to write archive JSON file: {filename}"
        ) from exc
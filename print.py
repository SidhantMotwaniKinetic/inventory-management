import sys
from pathlib import Path
import pandas as pd
import os
import tempfile
import subprocess

ROOT_DIR = Path(__file__).resolve().parent
STICKER_DIR = ROOT_DIR / "Inventronix"
BARTEND_EXE = r"C:\Program Files\Seagull\BarTender 2022\BarTend.exe"
PRINTER_NAME = "TSC TE244"
BTW_TEMPLATE = r"C:\Users\ems\Desktop\inventory-management\Inventronix\label.btw"

def print_label(id_value: str):
    # CSV with header 'id' matching your field name in BarTender
    csv_content = "id\n" + id_value + "\n"

    # Create temp CSV
    with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode="w", newline="") as f:
        data_path = f.name
        f.write(csv_content)

    try:
        # Call BarTender to print using that CSV
        subprocess.run([
            BARTEND_EXE,
            f"/AF={BTW_TEMPLATE}",   # your .btw
            f"/D={data_path}",       # data file
            "/P",                    # print
            f"/PRN={PRINTER_NAME}",  # printer
            "/X"                     # exit
        ], check=True)
    finally:
        os.remove(data_path)

def main():
    print("to print")
    print_label("IN01S0000K21-00$00$00$90000")
    print("dones")

if __name__ == "__main__":
    main()
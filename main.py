import sys
from pathlib import Path
import pandas as pd
import os
import re
import tempfile
import subprocess

import pyxlsb

##CHANGE ME
CONSOLIDATED_SHEET_NAME = "consolidated-18.xlsb"
SUPPLIER_ID_COLUMN = 0
REEL_SIZE_COLUMN = 4

ROOT_DIR = Path(__file__).resolve().parent
CONSOLIDATED_PATH = ROOT_DIR / "input" / CONSOLIDATED_SHEET_NAME
BARCODE_CSV_PATH = ROOT_DIR / "data" / "barcode-data-new-master.csv"
BARTEND_EXE = r"C:\Program Files\Seagull\BarTender 2022\BarTend.exe"
PRINTER_NAME = "TSC TE244"
BTW_TEMPLATE = r"C:\Users\ems\Desktop\inventory-management\Inventronix\label.btw"

def load_expected_quantities(consolidated_path: Path) -> dict:
    """
    Read the consolidated file (xlsx or xlsb) and build a dict:
        {supplier_id: {"total_quantity": int, "seen_quantity": int}}
    Assumptions:
    - Supplier ID is in column C
    - Quantity is in column J
    - First row is a header row
    """
    if not consolidated_path.exists():
        raise FileNotFoundError(f"Expected file not found: {consolidated_path}")

    # Determine file type and read accordingly
    if consolidated_path.suffix.lower() == ".xlsb":
        try:
            df = pd.read_excel(consolidated_path, engine="pyxlsb")
        except ImportError:
            raise ImportError("pyxlsb is required to read .xlsb files. Install it with: pip install pyxlsb")
    else:
        df = pd.read_excel(consolidated_path, engine="openpyxl")
    
    supplier_col_idx = SUPPLIER_ID_COLUMN
    qty_col_idx = REEL_SIZE_COLUMN

    if df.shape[1] <= max(supplier_col_idx, qty_col_idx):
        raise ValueError(
            "consolidated file does not appear to have the expected columns "
        )

    supplier_ids = df.iloc[:, supplier_col_idx]
    quantities = df.iloc[:, qty_col_idx]

    expected = {}
    for sid, qty in zip(supplier_ids, quantities):
        if pd.isna(sid) or pd.isna(qty):
            continue
        if sid == 'Grand Total':
            continue
        
        sid_str = str(sid).strip()
        try:
            qty_val = int(qty)
        except (TypeError, ValueError):
            continue

        if sid_str not in expected:
            expected[sid_str] = {"total_quantity": 0, "seen_quantity": 0}
        expected[sid_str]["total_quantity"] += qty_val

    return expected


def load_barcode_reference(barcode_csv_path: Path) -> dict:
    """
    Read barcode-data-master.csv and build a reference dict that only depends on:
        {supplier_id: {"kem_id": str}}

    Assumptions:
    - Column A: supplier id
    - Column B: kem id
    """
    if not barcode_csv_path.exists():
        raise FileNotFoundError(f"Reference CSV not found: {barcode_csv_path}")

    df = pd.read_csv(barcode_csv_path)

    if df.shape[1] < 2:
        raise ValueError(
            "barcode-data-master.csv does not appear to have at least 2 columns "
            "(A-B) for supplier id and kem id."
        )

    supplier_col_idx = 0
    kem_col_idx = 1

    supplier_ids = df.iloc[:, supplier_col_idx]
    kem_ids = df.iloc[:, kem_col_idx]

    reference = {}
    for sid, kem in zip(supplier_ids, kem_ids):
        if pd.isna(sid):
            continue
        sid_str = str(sid).strip()
        kem_str = "" if pd.isna(kem) else str(kem).strip()

        # If there are duplicates, last one will win. This can be adjusted if needed.
        reference[sid_str] = {
            "kem_id": kem_str,
        }

    return reference

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

def print_sticker(sticker: str):
    """Placeholder for printing sticker information."""
    if sticker:
        print(f" Sticker: {sticker}")
        print_label(sticker)
        print("Done Printing")
    else:
        print(" Erorr: No sticker specified.")

def generate_sticker_string(kem_id: str, reel_size: int) -> str:
    result = f"{kem_id}$00$00${reel_size}"
    print(f"Sticker: {result}")
    return result

def interactive_scan(expected: dict, reference: dict) -> None:
    """
    Interactive loop:
    - User enters supplier IDs (barcodes) and reel sizes one by one.
    - For each scan, print sticker info and update seen_quantity.
    - Enforce that seen_quantity never exceeds total_quantity for that ID.
    - Exit when user inputs an empty string or 'done' / 'exit' / 'quit' for the ID.
    """
    print("\n=== Inbound Inventory Scanning ===")
    print("Press Enter on a blank line or type 'done' when finished.\n")

    while True:
        #Fetch supplier ID from user input 
        try:
            scanned_id = input("Scan / enter supplier ID: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nStopping scan.")
            break
        if scanned_id == "" or scanned_id.lower() in {"done", "exit", "quit"}:
            break

        #clean out scanner noise (ANSI escape codes) and extract ID
        cleaned_id = re.sub(r"\x1b\[[0-9;]*[A-Za-z~]", "", scanned_id).strip()
        if not cleaned_id:
            print("  [ERROR] Supplier ID could not be parsed; please rescan.")
            continue
        supplier_id = cleaned_id.split("#", 1)[0]
        if supplier_id != scanned_id:
            print(
                "  [INFO] Cleaned supplier ID input. "
                f"Original: '{scanned_id}' -> Parsed: {supplier_id}"
            )
        print(f"Supplier ID: {supplier_id}")

        #Fetch reel size from user input
        try:
            reel_size_str = input("Scan / enter reel size: ").strip()
            if not reel_size_str:
                print("  [ERROR] Reel size cannot be empty.")
                continue

            try:
                reel_size = int(reel_size_str)
            except ValueError:
                digit_groups = re.findall(r"\d+", reel_size_str)
                if not digit_groups:
                    raise
                reel_size = int(max(digit_groups, key=len))
                print(
                    "  [INFO] Cleaned reel size input. "
                    f"Original: '{reel_size_str}' -> Parsed: {reel_size}"
                )

            print(f"Reel Size: {reel_size}")
        except (EOFError, KeyboardInterrupt):
            print("\nStopping scan.")
            break
        except ValueError:
            print(f"  [ERROR] Invalid reel size '{reel_size_str}'. Please enter a number.")
            continue

        #Check if supplier ID is in barcode reference file
        if supplier_id not in reference:
            print(f"  [ERROR] ID '{supplier_id}' not found in barcode reference file.")
            continue

        #Check if supplier ID is in expected consolidated list
        if supplier_id not in expected:
            print(
                f"  [ERROR] ID '{supplier_id}' was not present in expected consolidated list."
            )
            continue

        #Get kem ID from barcode reference file
        ref = reference[supplier_id]
        kem_id = ref["kem_id"]

        #Check if seen quantity exceeds total quantity
        rec = expected[supplier_id]
        new_seen = rec["seen_quantity"] + reel_size

        if new_seen > rec["total_quantity"]:
            print(
                f"  [ERROR] Scan would exceed expected quantity for '{supplier_id}'. "
                f"Expected: {rec['total_quantity']}, Currently seen: {rec['seen_quantity']}, "
                f"This reel: {reel_size}."
            )
            print(" ❌ This scan has been IGNORED. Please verify the item.\n")
            continue

        #Update seen quantity, remaining quantity
        rec["seen_quantity"] = new_seen
        remaining = rec["total_quantity"] - rec["seen_quantity"]
        print(
            f"  Updated seen quantity for '{supplier_id}': {rec['seen_quantity']} / {rec['total_quantity']}"
        )
        if remaining > 0:
            print(f"  Remaining quantity to scan for this ID: {remaining}")
        else:
            print("  ✅ This ID is now fully matched.")

        #Generate sticker string and print
        sticker_location = generate_sticker_string(kem_id, reel_size)
        print_sticker(sticker_location)
        print()

def final_reconciliation(expected: dict) -> None:
    """
    After scanning is done, compare total_quantity vs seen_quantity
    for each supplier ID and report a summary similar to the live scan logic:
    - Fully matched (seen == expected)
    - Partially matched (0 < seen < expected)
    - Not seen at all (seen == 0 < expected)
    """
    print("\n=== Final Reconciliation ===")
    fully_matched = []
    partial = []
    not_seen = []

    for sid, rec in expected.items():
        total_q = rec["total_quantity"]
        seen_q = rec["seen_quantity"]

        # Ignore any weird rows with non-positive expected quantity
        if total_q <= 0:
            continue

        if seen_q >= total_q:
            fully_matched.append((sid, total_q, seen_q))
        elif seen_q > 0:
            partial.append((sid, total_q, seen_q, total_q - seen_q))
        else:
            not_seen.append((sid, total_q))

    if not partial and not not_seen:
        print("All supplier IDs are fully matched. ✅ 🟢")
        return

    if fully_matched:
        print(f"🟢 Fully matched IDs ({len(fully_matched)}):")
        for sid, total_q, seen_q in fully_matched:
            print(f"  {sid}: {seen_q} / {total_q}")
        print()

    if partial:
        print("🟡 Partially matched IDs (some quantity still remaining):")
        print(f"{'Supplier ID':<20}{'Expected':>10}{'Seen':>10}{'Missing':>12}")
        print("-" * 52)
        for sid, total_q, seen_q, diff in partial:
            print(f"{sid:<20}{total_q:>10}{seen_q:>10}{diff:>12}")
        print()

    if not_seen:
        print("🔴 IDs with expected quantity but no scans at all:")
        print(f"{'Supplier ID':<20}{'Expected':>10}")
        print("-" * 32)
        for sid, total_q in not_seen:
            print(f"{sid:<20}{total_q:>10}")

def main() -> None:
    #parse input sheet and reference data
    try:
        expected = load_expected_quantities(CONSOLIDATED_PATH)
        reference = load_barcode_reference(BARCODE_CSV_PATH)

        print(expected)
        print("\n\n")
        # print(reference)
        # print("\n\n")

    except Exception as e:
        print(f"Failed to initialize data: {e}", file=sys.stderr)
        sys.exit(1)

    if not expected:
        print("No valid expected quantities found in consolidated.xlsx.", file=sys.stderr)
        sys.exit(1)

    if not reference:
        print("No valid barcode reference data found in barcode-data-new-master.csv.", file=sys.stderr)
        sys.exit(1)

    interactive_scan(expected, reference)
    final_reconciliation(expected)


if __name__ == "__main__":
    main()

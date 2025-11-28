import sys
from pathlib import Path
import pandas as pd
import os
import subprocess

CONSOLIDATED_SHEET_NAME = "CONSOLIDATED 27.11.2025.xlsx"
ROOT_DIR = Path(__file__).resolve().parent
CONSOLIDATED_PATH = ROOT_DIR / "input" / CONSOLIDATED_SHEET_NAME
BARCODE_CSV_PATH = ROOT_DIR / "data" / "barcode-data-master.csv"
STICKER_DIR = ROOT_DIR / "Inventronix"
BARTEND_EXE = r"C:\Program Files\Seagull\BarTender 2022\BarTend.exe"
PRINTER_NAME = "TSC TE244"
BTW_TEMPLATE = r"C:\Users\ems\Desktop\inventory-management\Inventronix\label.btw"

def print_btw_silent(btw_path: str):
    cmd = [
        BARTEND_EXE,
        f"/AF={btw_path}",        # BarTender format file
        "/P",                     # Print
        f"/PRN={PRINTER_NAME}",   # Target printer
        "/X"                      # Exit when done
    ]
    subprocess.run(cmd, check=True)

def print_sticker(sticker_location: str) -> None:
    print(f"Printing Sticker: {sticker_location}")
    try:
        base_folder = r"C:\Users\ems\Desktop\inventory-management\Inventronix"
        file_path = rf"{base_folder}\{sticker_location}"   # insert variable here
        print_btw_silent(file_path)
        print(f"Done Printing Sticker: {sticker_location}")
    except Exception as e:
        print(f"Error printing sticker: {e}")
    return

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

def print_custom_sticker(id: str) -> None:
    print(f"Printing Sticker: {sticker_location}")
    try:
        print_label(id)
        print(f"Done Printing Sticker: {sticker_location}")
    except Exception as e:
        print(f"Error printing sticker: {e}")
    return


def load_expected_quantities(consolidated_path: Path) -> dict:
    """
    Read the consolidated.xlsx file and build a dict:
        {supplier_id: {"total_quantity": int, "seen_quantity": int}}

    Assumptions:
    - Supplier ID is in column C
    - Quantity is in column J
    - First row is a header row
    """
    if not consolidated_path.exists():
        raise FileNotFoundError(f"Expected file not found: {consolidated_path}")

    # Read the whole sheet; default first sheet
    # Use column positions (0-based index: C -> 2, J -> 9)
    df = pd.read_excel(consolidated_path, engine="openpyxl")
    supplier_col_idx = 2
    qty_col_idx = 9

    if df.shape[1] <= max(supplier_col_idx, qty_col_idx):
        raise ValueError(
            "consolidated.xlsx does not appear to have the expected columns "
            "(need at least C and J)."
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
    Read barcode-data-master.csv and build a reference dict:
        {supplier_id: {"kem_id": str, "sticker_location": str, "reel_quantity": int}}

    Assumptions:
    - Column A: supplier id
    - Column B: kem id
    - Column C: sticker locations
    - Column D: reel quantity
    - First row is a header row (ignored automatically by pandas)
    """
    if not barcode_csv_path.exists():
        raise FileNotFoundError(f"Reference CSV not found: {barcode_csv_path}")

    df = pd.read_csv(barcode_csv_path)

    if df.shape[1] < 4:
        raise ValueError(
            "barcode-data-master.csv does not appear to have at least 4 columns "
            "(A-D) for supplier id, kem id, sticker location, reel quantity."
        )

    supplier_col_idx = 0
    kem_col_idx = 1
    sticker_col_idx = 2
    reel_qty_col_idx = 3

    supplier_ids = df.iloc[:, supplier_col_idx]
    kem_ids = df.iloc[:, kem_col_idx]
    stickers = df.iloc[:, sticker_col_idx]
    reel_qtys = df.iloc[:, reel_qty_col_idx]

    reference = {}
    for sid, kem, sticker, rqty in zip(
        supplier_ids, kem_ids, stickers, reel_qtys
    ):
        if pd.isna(sid):
            continue
        sid_str = str(sid).strip()
        kem_str = "" if pd.isna(kem) else str(kem).strip()
        sticker_str = "" if pd.isna(sticker) else str(sticker).strip()
        try:
            rqty_val = int(rqty)
        except (TypeError, ValueError):
            continue

        # If there are duplicates, last one will win. This can be adjusted if needed.
        reference[sid_str] = {
            "kem_id": kem_str,
            "sticker_location": sticker_str,
            "reel_quantity": rqty_val,
        }

    return reference


def interactive_scan(expected: dict, reference: dict) -> None:
    """
    Interactive loop:
    - User enters supplier IDs (barcodes) one by one.
    - For each scan, print sticker info and update seen_quantity based on reel quantity.
    - Enforce that seen_quantity never exceeds total_quantity for that ID.
    - Exit when user inputs an empty string or 'done' / 'exit' / 'quit'.
    """
    print("\n=== Inbound Inventory Scanning ===")
    print("Press Enter on a blank line or type 'done' when finished.\n")

    while True:
        try:
            scanned = input("Scan / enter supplier ID: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nStopping scan.")
            break

        if scanned == "" or scanned.lower() in {"done", "exit", "quit"}:
            break

        # supplier_id = scanned
        # for inventronics
        start_index = 0
        end_index = scanned.index('#')
        supplier_id = scanned[start_index:end_index]

        if supplier_id not in reference:
            print(f"  [ERROR] ID '{supplier_id}' not found in barcode reference file.")
            continue

        if supplier_id not in expected:
            print(
                f"  [ERROR] ID '{supplier_id}' was not present in expected consolidated list."
            )
            continue

        ref = reference[supplier_id]
        reel_qty = ref["reel_quantity"]
        sticker_location = ref["sticker_location"]
        kem_id = ref["kem_id"]

        rec = expected[supplier_id]
        new_seen = rec["seen_quantity"] + reel_qty

        if new_seen > rec["total_quantity"]:
            print(
                f"  [ERROR] Scan would exceed expected quantity for '{supplier_id}'. "
                f"Expected: {rec['total_quantity']}, Currently seen: {rec['seen_quantity']}, "
                f"This reel: {reel_qty}."
            )
            print(" ❌ This scan has been IGNORED. Please verify the item.\n")
            continue

        rec["seen_quantity"] = new_seen
        remaining = rec["total_quantity"] - rec["seen_quantity"]
        print(
            f"  Updated seen quantity for '{supplier_id}': {rec['seen_quantity']} / {rec['total_quantity']}"
        )
        if remaining > 0:
            print(f"  Remaining reels to scan for this ID: {remaining/reel_qty}")
        else:
            print("  ✅ This ID is now fully matched.")

        # print_sticker(sticker_location)
        print_custom_sticker(kem_id)

        print()


def final_reconciliation(expected: dict) -> None:
    """
    After scanning is done, compare total_quantity vs seen_quantity
    for each supplier ID and report discrepancies.
    """
    print("\n=== Final Reconciliation ===")
    missing = []

    for sid, rec in expected.items():
        total_q = rec["total_quantity"]
        seen_q = rec["seen_quantity"]
        if seen_q < total_q:
            missing.append((sid, total_q, seen_q, total_q - seen_q))

    if not missing:
        print("All supplier IDs are fully matched. ✅")
        return

    print("The following supplier IDs have remaining quantities:")
    print(f"{'Supplier ID':<20}{'Expected':>10}{'Seen':>10}{'Missing':>12}")
    print("-" * 52)
    for sid, total_q, seen_q, diff in missing:
        print(f"{sid:<20}{total_q:>10}{seen_q:>10}{diff:>12}")


def main() -> None:
    #parse input sheet and reference data
    try:
        expected = load_expected_quantities(CONSOLIDATED_PATH)
        reference = load_barcode_reference(BARCODE_CSV_PATH)

    except Exception as e:
        print(f"Failed to initialize data: {e}", file=sys.stderr)
        sys.exit(1)

    if not expected:
        print("No valid expected quantities found in consolidated.xlsx.", file=sys.stderr)
        sys.exit(1)

    if not reference:
        print("No valid barcode reference data found in barcode-data-master.csv.", file=sys.stderr)
        sys.exit(1)

    interactive_scan(expected, reference)
    final_reconciliation(expected)


if __name__ == "__main__":
    main()

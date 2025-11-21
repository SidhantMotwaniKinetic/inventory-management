# import subprocess

def extract_value(scanned_value):
    # Inventronics use case
    if '~' in scanned_value and '-' in scanned_value:
        start_index = scanned_value.index('~') + 1
        end_index = scanned_value.index('-', start_index)
        filename = "IN" + scanned_value[start_index:end_index] + ".bw"
        return filename
    else:
        print("Scanned input did not match expected pattern.")
    return None

def print_barcode_file(filename):
    print(filename)
    return
    # Example for macOS using 'lp' command, adjust for Windows (e.g., use 'print') or Linux
    # try:
    #     subprocess.run(['lp', filename], check=True)
    #     print(f"Sent file {filename} to printer.")
    # except Exception as e:
    #     print(f"Error printing {filename}: {e}")

def main():
    print("Scan barcodes. Press Ctrl+C to exit.")
    while True:
        scanned_value = input()
        value = extract_value(scanned_value)
        if value:
            filename = value
            print_barcode_file(filename)
        else:
            print("Scanned input did not match expected pattern.")

if __name__ == "__main__":
    main()

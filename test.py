import os
import subprocess
sticker_location = ""
print(f"Printing Sticker: {sticker_location}")
try:
    # file_path = STICKER_DIR + sticker_location
    # file_path = "01S1000110-00_1206_100R_1%.btw"
    # os.startfile(file_path)

    # btw_file = r"01S1000110-00_1206_100R_1%.btw"
    # print("inding print")
    # bartend_path = r"C:\Program Files\Seagull\BarTender\bartend.exe"  # Adjust if needed
    # subprocess.run([bartend_path, "/F", btw_file, "/P", "/X"])
    
    file_path = r"01S1000110-00_1206_100R_1%.btw"
    printer_nme = "TSC TE244"
    port = "USB001:"
    with open(file_path, 'rb') as f:
        data = f.read()
    with open(port, 'wb') as printer:
        printer.write(data)

    print(f"Done Printing Sticker: {sticker_location}")
except Exception as e:
    print(f"Error printing sticker: {e}")
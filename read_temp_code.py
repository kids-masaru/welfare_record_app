import os

try:
    with open("temp_old_main.py", "r", encoding="utf-16-le") as f:
        content = f.read()
        
    start_idx = content.find("def read_excel_monitoring_data")
    start_idx = content.find("def fill_excel")
    if start_idx != -1:
        # Write to file
        with open("extracted_functions.txt", "w", encoding="utf-8") as out:
            out.write(content[start_idx:start_idx+5000])
        print("Functions extracted to extracted_functions.txt")
    else:
        print("Neither function found.")

except Exception as e:
    print(e)

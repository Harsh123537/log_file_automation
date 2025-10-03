import re
import pandas as pd
import os
from datetime import datetime

# Folder containing log files
log_folder = r"C:\Users\M685200\Documents\log files"
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file = f"Web Service Log Monitoring_Add_column_{current_time}.xlsx"

# Regex pattern for timestamp
timestamp_pattern = r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})"

all_data = []

# Loop through each log file in folder
for log_file in os.listdir(log_folder):
    if log_file.endswith(".log") or log_file.endswith(".LOG"):   # only process .log files
        file_path = os.path.join(log_folder, log_file)

        data = []
        current_timestamp = None
        current_exception = None
        capture_message = False
        message_lines = [] 
        pre_exception_context=None 

        with open(file_path, "r", encoding="latin-1") as f:
            for line in f:
                line = line.rstrip()

                # Capture timestamp
                ts = re.match(timestamp_pattern, line)
                if ts:
                    current_timestamp = ts.group(1)

                    if " INFO " in line:
                        log_level = "INFO"
                    elif " ERROR " in line:
                        log_level = "ERROR"
                    else:
                        log_level = None

                    if log_level == "ERROR" and "Exception Name:" not in line and "Message:" not in line:
                        custom_msg = line.split(log_level, 1)[1].strip()
                        data.append([
                            log_file,
                            current_timestamp,
                            "",  # Exception Name blank
                            "",  # Error Message blank
                            log_level,  # INFO/ERROR column
                            custom_msg  # Custom Message
                        ])

                    # Extract context before "Exception Name:" if present
                    if "Exception Name:" in line:
                        pre_exception_context = line.split(log_level, 1)[1].split("Exception Name:")[0].strip()
                    else:
                        pre_exception_context = None

                # Detect Exception Name block
                if "Exception Name:" in line:
                    current_exception = None
                    continue

                # The line immediately after "Exception Name:" contains the exception
                if current_exception is None and line.strip():
                    current_exception = line.strip()
                    continue

                # Detect start of Message block
                if line.strip() == "Message:":
                    capture_message = True
                    message_lines = []  # reset for new message
                    continue

                # Capture multi-line error message
                if capture_message:
                    if line.startswith("Stacktrace:"):   # stop ONLY on Stacktrace
                        if message_lines:
                            full_message = "\n".join(message_lines).strip()
                            data.append([log_file, current_timestamp, current_exception, full_message,log_level,pre_exception_context])
                        capture_message = False
                        continue

                    if line.strip():  # skip empty lines
                        message_lines.append(line.strip())

        # Handle case where file ends while still capturing
        if capture_message and message_lines:
            full_message = "\n".join(message_lines).strip()
            data.append([log_file, current_timestamp, current_exception, full_message,log_level,pre_exception_context])

        all_data.extend(data)

# Convert all collected data to DataFrame
df = pd.DataFrame(all_data, columns=["Log File Name", "Timestamp", "Exception Name", "Error Message","INFO/ERROR","Custome Message"])

# Excel export with chunking to avoid row limit
EXCEL_MAX_ROWS = 1048576
chunk_size = EXCEL_MAX_ROWS - 1  # leave 1 row for header

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    total_rows = len(df)
    sheet_number = 1

    for start_row in range(0, total_rows, chunk_size):
        end_row = start_row + chunk_size
        df_chunk = df.iloc[start_row:end_row]

        sheet_name = f"Part_{sheet_number}"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]

        df_chunk.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
        sheet_number += 1

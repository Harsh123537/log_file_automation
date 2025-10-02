import re
import pandas as pd
import os
from datetime import datetime

# Folder containing log files
log_folder = r"C:\Users\Harsh Agrawal\Documents\log files"
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file = f"Web Service Log Monitoring_{current_time}.xlsx"

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
        message_lines = []  # store multi-line messages

        with open(file_path, "r", encoding="latin-1") as f:
            for line in f:
                line = line.rstrip()

                # Capture timestamp
                ts = re.match(timestamp_pattern, line)
                if ts:
                    current_timestamp = ts.group(1)

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
                            data.append([log_file, current_timestamp, current_exception, full_message])
                        capture_message = False
                        continue

                    if line.strip():  # skip empty lines
                        message_lines.append(line.strip())

        # Handle case where file ends while still capturing
        if capture_message and message_lines:
            full_message = "\n".join(message_lines).strip()
            data.append([log_file, current_timestamp, current_exception, full_message])

        all_data.extend(data)

# Convert all collected data to DataFrame
df = pd.DataFrame(all_data, columns=["Log File Name", "Timestamp", "Exception Name", "Error Message"])

# Save everything to Excel
df.to_excel(output_file, index=False)

# print(f"✅ Extracted data saved to {output_file}")

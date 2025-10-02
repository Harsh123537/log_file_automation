import re
import pandas as pd

log_file = "ProdQualityComplaint.log"
output_file = "ProdQualityComplaint.xlsx"

# patterns
timestamp_pattern = r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3})"
pqccallcenter_pattern = r"PQCCallCenterID:\s*([A-Z]{3}\d{2}-\d{6})"   # e.g. USA25-015049
new_query_pqc_pattern = r"New Query as received from Call Center:.*\b([A-Z]{3}\d{2}-\d{6})\b"
# product_pattern = r"ProductName:\s*(.+)"
gqc_pr_pattern = r"GQC PR# - (\d+)"   # only accept PR IDs of this exact form

data = []

# state
current_timestamp = None
current_pqccallcenter = None
# current_product = None
current_prid = None
capture_message = False

with open(log_file, "r", encoding="latin-1") as f:
    for line in f:
        line = line.rstrip()

        # timestamp (keep for rows)
        ts = re.match(timestamp_pattern, line)
        if ts:
            current_timestamp = ts.group(1)

        # explicit PQCCallCenterID: header (reset PR ID for a new record)
        m = re.search(pqccallcenter_pattern, line)
        if m:
            current_pqccallcenter = m.group(1)
            current_prid = None  # IMPORTANT: reset PR ID when a new PQC record starts
            continue

        # sometimes PQC appears embedded in the "New Query..." DEBUG line
        m = re.search(new_query_pqc_pattern, line)
        if m:
            current_pqccallcenter = m.group(1)
            current_prid = None
            continue

        # # product name (captured for context)
        # m = re.search(product_pattern, line)
        # if m:
        #     current_product = m.group(1).strip()
        #     continue

        # Only set PR ID when the exact "GQC PR# - <digits>" appears
        m = re.search(gqc_pr_pattern, line)
        if m:
            current_prid = m.group(1)
            continue

        # detect Message: line
        if line.strip() == "Message:":
            capture_message = True
            continue

        # capture first non-empty line after Message:, skip blank lines, stop before Stacktrace: or new timestamp
        if capture_message:
            # skip blank lines between "Message:" and the actual message
            if not line.strip():
                continue

            # if we hit Stacktrace: or a new timestamp we stop without adding
            if line.startswith("Stacktrace:") or re.match(timestamp_pattern, line):
                capture_message = False
                continue

            # this is the actual message (first non-empty line after Message:)
            message = line.strip()
            data.append({
                "Timestamp": current_timestamp,
                "PQCCallCenterID": current_pqccallcenter or "",
                # "ProductName": current_product or "",
                "PR ID": current_prid or "",   # will be empty if no "GQC PR#" seen for this block
                "Error Message": message
            })
            # done capturing this message
            capture_message = False
            continue

# save to excel
df = pd.DataFrame(data)
df.to_excel(output_file, index=False)
# print(f"Extracted {len(df)} error(s) -> {output_file}")
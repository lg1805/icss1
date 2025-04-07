from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from datetime import datetime
import xlsxwriter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# RPN file
RPN_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx"
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

def extract_component(observation):
    if pd.notna(observation):
        for component in known_components:
            if str(component).lower() in observation.lower():
                return component
    return "Unknown"

def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        severity = int(row["Severity (S)"].values[0])
        occurrence = int(row["Occurrence (O)"].values[0])
        detection = int(row["Detection (D)"].values[0])
        return severity, occurrence, detection
    return 1, 1, 10

def determine_priority(rpn):
    if rpn >= 200:
        return "High"
    elif rpn >= 100:
        return "Moderate"
    else:
        return "Low"

def month_str_to_num(month_hint):
    month_map = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04",
        "may": "05", "jun": "06", "jul": "07", "aug": "08",
        "sep": "09", "oct": "10", "nov": "11", "dec": "12"
    }
    return month_map.get(month_hint.lower(), None)

def format_creation_date(date_str, month_hint):
    target_month = month_str_to_num(month_hint)
    if not target_month:
        return None, None

    try:
        date_str = str(date_str).strip()
        dt = pd.to_datetime(date_str, dayfirst=True, errors='coerce')

        if pd.notna(dt):
            dd, mm, yyyy = dt.day, dt.month, dt.year

            if str(dd).zfill(2) == "01" and str(mm).zfill(2) == "01":
                dd = int(target_month)

            formatted_date = f"{str(dd).zfill(2)}/{target_month}/{yyyy}"
            elapsed_days = (datetime.now() - dt).days
            return formatted_date, elapsed_days

    except Exception:
        return None, None

    return None, None

@app.route('/')
def index():
    return render_template('frontNEW.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No complaint_file part", 400

    file = request.files['complaint_file']
    if file.filename == '':
        return "No selected file", 400

    month_hint = request.form.get('month_hint', 'default')

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            return f"Error reading file: {e}", 400

        if 'Observation' not in df.columns or 'Creation Date' not in df.columns or 'Incident no' not in df.columns:
            return "Required columns missing", 400

        # Format date and calculate elapsed
        formatted_dates = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
        df['Creation Date'] = formatted_dates.apply(lambda x: x[0])
        df['Elapsed Days'] = formatted_dates.apply(lambda x: x[1])
        df['Elapsed Days'] = df['Elapsed Days'].fillna(-1).astype(int)  # avoid NaN issues

        def get_color(elapsed):
            if elapsed == 1:
                return '#ADD8E6'  # Light Blue
            elif elapsed == 2:
                return '#FFFF00'  # Yellow
            elif elapsed == 3:
                return '#FF1493'  # Pink
            elif elapsed > 3:
                return '#FF0000'  # Red
            else:
                return None

        # RPN-related logic
        df["Component"] = df["Observation"].apply(extract_component)
        df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = df["Component"].apply(lambda comp: pd.Series(get_rpn_values(comp)))
        df["RPN"] = df["Severity (S)"] * df["Occurrence (O)"] * df["Detection (D)"]
        df["Priority"] = df["RPN"].apply(determine_priority)

        # Split SPN and non-SPN
        spn_df = df[df["Observation"].str.contains("spn", case=False, na=False)]
        non_spn_df = df[~df["Observation"].str.contains("spn", case=False, na=False)]

        # Sort by priority
        priority_order = {"High": 1, "Moderate": 2, "Low": 3}
        spn_df = spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))
        non_spn_df = non_spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))

        # Final path
        processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)

        with pd.ExcelWriter(processed_filepath, engine='xlsxwriter') as writer:
            for sheet_name, sheet_df in zip(["SPN", "Non-SPN"], [spn_df, non_spn_df]):
                sheet_df = sheet_df.fillna('')  # Important to prevent NaN write errors
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                # Format for coloring
                green_fmt = workbook.add_format({'bg_color': '#006400'})  # Green for closed
                formats = {
                    1: workbook.add_format({'bg_color': '#ADD8E6'}),
                    2: workbook.add_format({'bg_color': '#FFFF00'}),
                    3: workbook.add_format({'bg_color': '#FF1493'}),
                    4: workbook.add_format({'bg_color': '#FF0000'}),
                }

                for idx, row_idx in enumerate(sheet_df.index):
                    elapsed = int(sheet_df.loc[row_idx, 'Elapsed Days'])
                    incident_status = str(sheet_df.loc[row_idx, "Incident Status"]).lower()

                    # Color Incident no if elapsed days match
                    if elapsed >= 1 and elapsed <= 3:
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident no"), sheet_df.loc[row_idx, "Incident no"], formats[elapsed])
                    elif elapsed > 3:
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident no"), sheet_df.loc[row_idx, "Incident no"], formats[4])

                    # Green if closed
                    if "closed" in incident_status or "complete" in incident_status:
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident Status"), sheet_df.loc[row_idx, "Incident Status"], green_fmt)

        return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

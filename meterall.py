import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

def to_excel_with_bold(df):
    out = BytesIO()
    wb = Workbook()
    ws = wb.active
    for i, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        ws.append(row)
        if i == 0:
            for cell in ws[1]:
                cell.font = Font(bold=True)
    wb.save(out)
    return out.getvalue()

st.title("All‑Meters Δm³/day Generator")

# 1) Upload
uploaded = st.file_uploader("Upload raw meter readings Excel (.xlsx)", type="xlsx")
if not uploaded:
    st.stop()

# 2) Read & filter
raw = pd.read_excel(uploaded)
raw = raw.dropna(axis=1, how='all')
date_col = raw.columns[0]
raw[date_col] = pd.to_datetime(raw[date_col], errors='coerce')
meters = [c for c in raw.columns if "reading" in c.lower()]
df = raw[[date_col] + meters].dropna(subset=meters, how='all').sort_values(by=date_col)

# 3) Date inputs
min_d = df[date_col].dt.date.min()
max_d = df[date_col].dt.date.max()
start_input = st.date_input("Start Date", value=min_d, min_value=min_d, max_value=max_d)
end_input   = st.date_input("End Date",   value=max_d, min_value=min_d, max_value=max_d)
if start_input > end_input:
    st.error("Start must be ≤ End")
    st.stop()

# 4) plot_size = 1
plot_size = 1

# 5) Build calendar
calendar = pd.date_range(start=start_input, end=end_input, freq='D').date
out = pd.DataFrame({"Date": calendar})

# 6) Compute and align each meter
for m in meters:
    temp = df[[date_col, m]].copy()
    temp[m] = pd.to_numeric(temp[m], errors='coerce')
    temp = temp.dropna(subset=[m]).sort_values(by=date_col)
    temp["Date_only"] = temp[date_col].dt.date

    days  = temp[date_col].diff().dt.days.fillna(method='bfill').replace(0,1)
    delta = temp[m].diff().fillna(method='bfill')
    rate  = delta / days

    collapsed = pd.DataFrame({
        "Date": temp["Date_only"],
        "Rate": rate
    }).groupby("Date", as_index=False).first()

    uniq_dates = np.array(collapsed["Date"], dtype="datetime64[D]")
    uniq_rates = collapsed["Rate"].to_numpy()

    cal_arr = np.array(calendar, dtype="datetime64[D]")
    idxs = np.searchsorted(uniq_dates, cal_arr, side="left")

    # Safely build full_rates
    full_rates = np.zeros_like(cal_arr, dtype=float)
    mask = idxs < len(uniq_rates)
    full_rates[mask] = uniq_rates[idxs[mask]]

    out[f"{m}-Dm3/dspr"] = full_rates

# 7) Preview & Download
st.dataframe(out.head(10))
excel_bytes = to_excel_with_bold(out)
st.download_button(
    "Download All‑Meters Δm³/day",
    data=excel_bytes,
    file_name="All_Meters_Delta_Per_Day.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

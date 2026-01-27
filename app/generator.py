import pandas as pd
from io import BytesIO
from typing import Tuple, Optional
from datetime import datetime
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
import logging


class NoDataInRange(Exception):
    """Raised when no rows remain after the date filter."""
    pass


def generate_clawback_report(
    input1_bytes: bytes,
    input2_bytes: bytes,
    start_date_str: str,
    end_date_str: str,
    report_title: Optional[str] = None
) -> Tuple[bytes, str]:
    """
    In-memory version of your CLI script. Mirrors the first script's logic/formatting.

    Parameters
    ----------
    input1_bytes : bytes
        Raw Excel content for the commissions data (same as --input1 in CLI).
    input2_bytes : bytes
        Raw Excel content for the team mapping (same as --input2 in CLI).
    start_date_str : str
        Inclusive start date for filtering 'Start Date' (e.g., '2024-04-01').
    end_date_str : str
        Inclusive end date for filtering 'Start Date' (e.g., '2025-03-31').
    report_title : Optional[str]
        Title displayed above each Manager Report section. If None, derived from dates.

    Returns
    -------
    (excel_bytes, filename) : Tuple[bytes, str]
        The generated Excel file as bytes, and a suggested filename.
    """

    # ------------------------------------------------------------
    # 1) Read Excel and cleanup  (parity with CLI)
    # ------------------------------------------------------------
    # Raw input: skip the first row, all text

    logging.info("Process started.") 
    
    logging.info("Reading input 1st Excel files.")
    df_raw = pd.read_excel(BytesIO(input1_bytes), dtype=str, skiprows=1, engine="openpyxl")
    print(df_raw.columns)  # parity with CLI (debug prints)
    df_raw.columns = df_raw.columns.str.strip()

   
    
    print("=== EXACT COLUMN NAMES ===")  # parity with CLI (debug prints)
    for i, col in enumerate(df_raw.columns):
        print(f"{i}: {repr(col)}")
    print("===========================")

    # Filter remarks exactly as CLI
    df_raw = df_raw[df_raw['Remarks'].isin(['Adj', 'Claw', 'New'])]

    # Adviser Name
    df_raw['Adviser Name'] = df_raw['Adviser Forename'] + ' ' + df_raw['Adviser Surname']
    df_raw['Adviser Name'] = df_raw['Adviser Name'].str.replace(' 1', '', regex=True)

    # APE cleanup (mirror CLI)
    df_raw['APE'] = df_raw['APE'].str.replace(',', '', regex=True)
    df_raw['APE'] = df_raw['APE'].str.replace(' -   ', '0')
    df_raw['APE'] = df_raw['APE'].str.replace('#DIV/0!', '0', regex=True)
    df_raw['APE'] = pd.to_numeric(df_raw['APE']).dropna()
    df_raw = df_raw[df_raw['APE'].abs() > 1]

    # Dates + filter (inclusive)
    df_raw['Paid Date'] = pd.to_datetime(df_raw['Paid Date'], errors='coerce')
    df_raw['Start Date'] = pd.to_datetime(df_raw['Start Date'], errors='coerce')

    start_date = pd.to_datetime(start_date_str, errors="raise")
    end_date   = pd.to_datetime(end_date_str, errors="raise")

    df_raw = df_raw[(df_raw['Start Date'] >= start_date) & (df_raw['Start Date'] <= end_date)]
    if df_raw.empty:
        formatted_start = start_date.strftime('%d-%B-%Y')
        formatted_end = end_date.strftime('%d-%B-%Y')
        raise NoDataInRange(f"No data found between {formatted_start} and {formatted_end}")

    print(df_raw)  # parity with CLI (debug prints)

    # ------------------------------------------------------------
    # 2) NTU policies (zero-payment)
    # ------------------------------------------------------------
    policy_totals = (
        df_raw
        .groupby('Policy Number')['APE']
        .sum()
        .round(2)
        .reset_index()
    )

    zero_policy_totals = policy_totals[policy_totals['APE'] == 0]
    ntu_policies = zero_policy_totals['Policy Number']

    filtered_df = df_raw[df_raw['Policy Number'].isin(ntu_policies)]
    main_df = (
        filtered_df
        .groupby('Adviser Name')['Policy Number']
        .nunique()
        .reset_index(name='NTU Count')
    )

    positive_ntu = filtered_df[filtered_df['APE'] > 0]
    ntu_payment = (
        positive_ntu
        .groupby('Adviser Name')['APE']
        .sum()
        .round(2)
        .reset_index(name='NTU APE')
    )

    main_df = main_df.merge(ntu_payment, on='Adviser Name', how='left')

    # ------------------------------------------------------------
    # 3) Clawbacks
    # ------------------------------------------------------------
    df_raw_no_ntu = df_raw[~df_raw['Policy Number'].isin(ntu_policies)]
    df_raw_no_ntu['Clawback'] = df_raw_no_ntu['Remarks'].str.contains('Claw|Adj', case=False, na=False).astype(int)

    clawback_count = (
        df_raw_no_ntu[df_raw_no_ntu['Clawback'] == 1]
        .groupby('Adviser Name')['Policy Number']
        .nunique()
        .reset_index(name='Clawback Count')
    )

    clawback_payment = (
        df_raw_no_ntu[(df_raw_no_ntu['Clawback'] == 1) & (df_raw_no_ntu['APE'] < 0)]
        .groupby('Adviser Name')['APE']
        .sum()
        .round(2)
        .reset_index(name='Clawback APE')
    )
    # Make APE positive (represents amount clawed back)
    clawback_payment['Clawback APE'] *= -1

    df_raw_clawback = clawback_count.merge(clawback_payment, on='Adviser Name', how='left')
    main_df = main_df.merge(df_raw_clawback, on='Adviser Name', how='outer')

    # ------------------------------------------------------------
    # 4) Raw issued business + counts (Remarks == 'New')
    # ------------------------------------------------------------
    df_raw_positive = df_raw[df_raw['Remarks'] == 'New']

    raw_total_payment = (
        df_raw_positive
        .groupby('Adviser Name')['APE']
        .sum()
        .round(2)
        .reset_index(name='APE')
    )

    raw_policy_count = (
        df_raw_positive
        .groupby('Adviser Name')['Policy Number']
        .nunique()
        .reset_index(name='Total Count')
    )

    df_raw_agg = raw_total_payment.merge(raw_policy_count, on='Adviser Name', how='left')
    main_df = main_df.merge(df_raw_agg, on='Adviser Name', how='outer').fillna(0)

    print(main_df[main_df['Adviser Name'] == 'Adrian Tabus'])  # parity with CLI (debug prints)

    # ------------------------------------------------------------
    # 5) Calculate percentages
    # ------------------------------------------------------------
    main_df['NTU %'] = (
        main_df['NTU Count'].fillna(0)
        / main_df['Total Count'].replace(0, float('nan'))
    ).round(4)

    main_df['Clawback %'] = (
        main_df['Clawback Count'].fillna(0)
        / main_df['Total Count'].replace(0, float('nan'))
    ).round(4)

    main_df['Total %'] = (main_df['NTU %'] + main_df['Clawback %']).round(4)
    main_df.rename(columns={'Adviser Name': 'Adviser'}, inplace=True)

    # ------------------------------------------------------------
    # 6) Merge Team Group (mapping)
    # ------------------------------------------------------------
    
    logging.info("Reading input 2nd Excel files.")
    
    df_leave_manager = pd.read_excel(BytesIO(input2_bytes), dtype=str, engine="openpyxl")
    df_leave_manager.rename(columns={'NAME ': 'Adviser'}, inplace=True)
    df_leave_manager['Team Group'] = df_leave_manager['Team Group'].str.rstrip()

    main_df = main_df.merge(
        df_leave_manager[['Adviser', 'Team Group']],
        on='Adviser',
        how='left'
    )

    # Keep only advisers that exist in the mapping file
    valid_advisers = set(df_leave_manager['Adviser'].unique())
    main_df = main_df[main_df['Adviser'].isin(valid_advisers)]

    # Reorder columns
    cols = ['Adviser', 'Team Group'] + [
        col for col in main_df.columns
        if col not in ['Adviser', 'Team Group']
    ]
    main_df = main_df[cols]

    # ------------------------------------------------------------
    # 7) Final metrics and ordering
    # ------------------------------------------------------------
    main_df["IN Books $"] = main_df["APE"] - main_df["Clawback APE"]
    main_df["IN Books Count"] = (
        main_df["Total Count"] - (main_df["Clawback Count"] + main_df["NTU Count"])
    )

    desired_columns = [
        "Adviser",
        "Team Group",
        "APE",
        "Clawback APE",
        "NTU APE",
        "IN Books $",
        "Total Count",
        "Clawback Count",
        "NTU Count",
        "IN Books Count",
        "Clawback %",
        "NTU %",
        "Total %"
    ]
    main_df = main_df[desired_columns]
    main_df.rename(columns={
        'APE': 'Issued Business $',
        'Clawback APE': 'Clawback $',
        'NTU APE': 'NTU $',
        'Total Count': 'Issued Policies Count'
    }, inplace=True)

    # 8) Cleanup/scale
    main_df.fillna(0, inplace=True)
    main_df = main_df[main_df['Team Group'] != 'Left Adviser']

    percent_cols = ["Clawback %", "NTU %", "Total %"]
    main_df[percent_cols] *= 100

    # (Optional prints to mirror CLI; they don't affect output)
    sum_row = main_df.drop(columns=percent_cols).sum(numeric_only=True)
    avg_percent_row = main_df[percent_cols].mean()
    summary_row = pd.concat([sum_row, avg_percent_row])
    summary_df = pd.DataFrame(summary_row, columns=["Value"]).T

    print(summary_df)  # parity with CLI (debug prints)
    print(main_df['Issued Business $'].sum())  # parity with CLI (debug prints)

    # ------------------------------------------------------------
    # 9) Provider Pivot
    # ------------------------------------------------------------
    df_pivot = (
        df_raw[['Adviser Name', 'Provider', 'Policy Number']]
        .drop_duplicates()
        .merge(df_leave_manager[['Adviser', 'Team Group']],
               left_on='Adviser Name', right_on='Adviser', how='left')
    )

    df_pivot = (
        df_pivot
        .groupby(['Adviser', 'Team Group', 'Provider'])['Policy Number']
        .nunique()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Total and ordering
    df_pivot['Total Policies Sold'] = df_pivot.iloc[:, 2:].sum(axis=1)
    cols_pivot = df_pivot.columns.tolist()
    provider_cols = cols_pivot[2:-1]
    new_order = ['Adviser', 'Team Group', 'Total Policies Sold'] + provider_cols
    df_pivot = df_pivot[new_order]

    # % columns
    for col in provider_cols:
        df_pivot[f'{col} %'] = (df_pivot[col] / df_pivot['Total Policies Sold']) * 100
    pct_cols_gen = [f'{col} %' for col in provider_cols]
    if pct_cols_gen:
        df_pivot[pct_cols_gen] = df_pivot[pct_cols_gen].round(2)

    df_pivot = df_pivot[df_pivot['Team Group'] != 'Left Adviser']

    # ------------------------------------------------------------
    # 10) Manager Report assembly
    # ------------------------------------------------------------
    
    logging.info("Assembling Manager Report.")
    
    main_df_sorted = main_df.sort_values(['Team Group','Adviser'], na_position='last')
    grouped = main_df_sorted.groupby('Team Group', dropna=False)

    manager_rows = []
    ordered_columns = [col for col in main_df.columns if col != 'Team Group']
    first_group = True

    for team_group, grp_df in grouped:
        # 3 blank lines between sections (except before first)
        if not first_group:
            blank = {col: '' for col in ordered_columns}
            manager_rows.extend([blank.copy(), blank.copy(), blank.copy()])
        first_group = False

        # Headers row
        headers_row = {col: col for col in ordered_columns}
        manager_rows.append(headers_row)

        # Subtotals for this group
        sum_issued_business  = grp_df['Issued Business $'].sum()
        sum_clawback_dollars = grp_df['Clawback $'].sum()
        sum_ntu_dollars      = grp_df['NTU $'].sum()
        sum_in_books         = grp_df['IN Books $'].sum()
        sum_issued_policies  = grp_df['Issued Policies Count'].sum()
        sum_clawback_count   = grp_df['Clawback Count'].sum()
        sum_ntu_count        = grp_df['NTU Count'].sum()
        sum_in_books_count   = grp_df['IN Books Count'].sum()

        if sum_issued_policies == 0:
            team_clawback_pct = 0
            team_ntu_pct      = 0
        else:
            team_clawback_pct = sum_clawback_count / sum_issued_policies
            team_ntu_pct      = sum_ntu_count   / sum_issued_policies

        team_total_pct = team_clawback_pct + team_ntu_pct

        subtotal_row = {
            'Adviser':               f'Team {str(team_group).upper()}:',
            'Issued Business $':     sum_issued_business,
            'Clawback $':            sum_clawback_dollars,
            'NTU $':                 sum_ntu_dollars,
            'IN Books $':            sum_in_books,
            'Issued Policies Count': sum_issued_policies,
            'Clawback Count':        sum_clawback_count,
            'NTU Count':             sum_ntu_count,
            'IN Books Count':        sum_in_books_count,
            'Clawback %':            round(team_clawback_pct*100, 2),
            'NTU %':                 round(team_ntu_pct*100, 2),
            'Total %':               round(team_total_pct*100, 2)
        }
        manager_rows.append(subtotal_row)

        # Detail rows
        for _, detail_row in grp_df[ordered_columns].iterrows():
            manager_rows.append(detail_row.to_dict())

    manager_report_df = pd.DataFrame(manager_rows, columns=ordered_columns)

    # ------------------------------------------------------------
    # 11) Excel writing + formatting  (mirrors CLI)
    # ------------------------------------------------------------
    
    logging.info("Generating Excel report.")
    
    wb = Workbook()

    # Styles
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center')
    bold_border = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")
    blue_fill = PatternFill(start_color="B0C4DE", end_color="B0C4DE", fill_type="solid")
    UNIFORM_WIDTH = 18

    # Sheet 1: Clawback Report
    ws1 = wb.active
    ws1.title = "Clawback Report"
    for r in dataframe_to_rows(main_df, index=False, header=True):
        ws1.append(r)

    # Sheet 2: Business Mix
    ws2 = wb.create_sheet("Business Mix")
    for r in dataframe_to_rows(df_pivot, index=False, header=True):
        ws2.append(r)

    # Sheet 3: Manager Report
    ws3 = wb.create_sheet("Manager Report")
    for r in dataframe_to_rows(manager_report_df, index=False, header=False):
        ws3.append(r)

    # --- Manager Report styling block (same as CLI) ---
    row_idx = 1
    while row_idx <= ws3.max_row:
        adviser_cell = ws3.cell(row=row_idx, column=1).value

        if isinstance(adviser_cell, str) and adviser_cell.startswith("Adviser"):
            # Header row
            for col in range(1, ws3.max_column + 1):
                cell = ws3.cell(row=row_idx, column=col)
                cell.font = bold_font
                cell.border = bold_border
                cell.fill = light_blue_fill

            # Subtotal row (next row)
            for col in range(1, ws3.max_column + 1):
                cell = ws3.cell(row=row_idx + 1, column=col)
                cell.font = bold_font
                cell.border = bold_border

            row_idx += 3
        else:
            row_idx += 1

    # Auto-adjust column widths (CLI behavior before uniform widths are applied later)
    for col in ws3.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                val = str(cell.value) if cell.value is not None else ""
                if len(val) > max_length:
                    max_length = len(val)
            except Exception:
                pass
        ws3.column_dimensions[column].width = max_length + 2  # padding

    # Conditional fill for % columns on Manager Report (like CLI block)
    percent_columns = ['Clawback %', 'NTU %', 'Total %']
    percent_col_indices = [ordered_columns.index(col) + 1 for col in percent_columns if col in ordered_columns]

    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row):
        for col_idx in percent_col_indices:
            cell = row[col_idx - 1]
            try:
                val = float(cell.value)
                if 10 <= val < 15:
                    cell.fill = yellow_fill
                elif val >= 15:
                    cell.fill = red_fill
            except (TypeError, ValueError):
                continue

    # Insert report title and a blank row above each section
    if not report_title:
        report_title = f"Clawback Report {start_date.strftime('%b %Y')} to {end_date.strftime('%b %Y')}"
    row_idx = 1
    while row_idx <= ws3.max_row:
        first_cell = ws3.cell(row=row_idx, column=1).value
        # Detect the start of a section (headers row)
        if isinstance(first_cell, str) and first_cell == "Adviser":
            ws3.insert_rows(row_idx)
            ws3.cell(row=row_idx, column=1, value=report_title).font = bold_font
            row_idx += 2  # Skip past the inserted title + blank row
        row_idx += 1

    # --- Universal formatting function (as in CLI) ---
    def format_sheet(ws, apply_coloring=False):
        headers = [cell.value for cell in ws[1]]
        adviser_idx = next((i + 1 for i, val in enumerate(headers) if val == 'Adviser'), None)
        percent_col_idxs = [i + 1 for i, val in enumerate(headers) if isinstance(val, str) and "%" in val]

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for col_idx, cell in enumerate(row, 1):
                # Header styling
                if row[0].row == 1:
                    cell.font = bold_font
                    cell.alignment = center_align
                    cell.border = bold_border

                # Uniform borders
                if ws.title in ["Clawback Report", "Business Mix"]:
                    cell.border = bold_border

                # Bold Adviser column
                if adviser_idx is not None and col_idx == adviser_idx:
                    cell.font = bold_font

                # Center-align numbers
                if isinstance(cell.value, (int, float)):
                    cell.alignment = center_align

                # Conditional formatting to % columns
                if apply_coloring and col_idx in percent_col_idxs:
                    try:
                        val = float(cell.value)
                        if 10 <= val < 15:
                            cell.fill = yellow_fill
                        elif val >= 15:
                            cell.fill = red_fill
                    except (TypeError, ValueError):
                        pass

        # Set uniform column width
        for col in ws.columns:
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = UNIFORM_WIDTH

    # Apply formatting to all three sheets (same as CLI)
    format_sheet(ws1, apply_coloring=True)   # Clawback Report
    format_sheet(ws2, apply_coloring=False)  # Business Mix (no color)
    format_sheet(ws3, apply_coloring=True)   # Manager Report

    # --- TOTAL row on Clawback Report (exactly like CLI) ---
    headers = [cell.value for cell in ws1[1]]
    percent_cols = ['Clawback %', 'NTU %', 'Total %']

    # Blank row
    ws1.append([""] * len(headers))

    # Calculate totals
    data_rows = list(ws1.iter_rows(min_row=2, max_row=ws1.max_row - 1))  # exclude header and last blank row
    totals = {}
    for col_idx, header in enumerate(headers, 1):
        if header not in percent_cols:
            total = 0.0
            for row in data_rows:
                try:
                    val = row[col_idx - 1].value
                    total += float(val) if val is not None else 0.0
                except Exception:
                    pass
            totals[header] = total

    issued_count = totals.get('Issued Policies Count', 0)
    totals['Clawback %'] = round(totals.get('Clawback Count', 0) / issued_count * 100, 2) if issued_count else 0
    totals['NTU %']       = round(totals.get('NTU Count', 0) / issued_count * 100, 2) if issued_count else 0
    totals['Total %']     = totals['Clawback %'] + totals['NTU %']

    total_row_values = []
    for header in headers:
        total_row_values.append("TOTAL" if header == 'Adviser' else totals.get(header, 0))
    ws1.append(total_row_values)

    # Style total row
    last_row = ws1.max_row
    for col_idx in range(1, len(headers) + 1):
        cell = ws1.cell(row=last_row, column=col_idx)
        cell.font = bold_font
        cell.fill = blue_fill
        cell.border = bold_border
        cell.alignment = center_align

    # ------------------------------------------------------------
    # Serialize to bytes
    # ------------------------------------------------------------
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    current_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"clawback_report_{current_ts}.xlsx"
    return buf.read(), filename

    logging.info("Process completed.")

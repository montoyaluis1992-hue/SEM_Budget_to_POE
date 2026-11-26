#!/usr/bin/env python
# coding: utf-8

import os
import traceback
import numpy as np
import pandas as pd
import xlsxwriter


# =====================================================
# LOAD ASSETS & SPEND DATA
# =====================================================

def get_file(filename):
    return os.path.join('POE Report Assets', filename)

def load_inputs():
    asset_file_path = get_file('SA360_SEM_Report.csv')
    if not os.path.exists(asset_file_path):
        raise FileNotFoundError(f'File not found: {asset_file_path}')

    asset_file = pd.read_csv(asset_file_path, skiprows=2)
    print('Inputs Loaded')
    return asset_file

def load_inputs2():
    budget_file = 'SEM Budget Input.xlsx'
    daily_spend_file = 'SEM Daily Spend.xlsx'

    q_budget = pd.read_excel(budget_file, sheet_name='QB', skiprows=7)
    daily_spend = pd.read_excel(daily_spend_file, sheet_name='SEM Daily Spend SA360', skiprows=2)
    campaign_conditions = pd.read_excel(budget_file, sheet_name='Campaign_conditions')
    return q_budget, daily_spend, campaign_conditions


def main():
    asset_file = load_inputs()
    q_budget, daily_spend, campaign_conditions = load_inputs2()

    # ===============================
    # LOOKUP CREATIVES
    # ===============================
    sa360_lookup_cols = [
        'Campaign',
        'Headline 1',
        'Headline 2',
        'Description',
        'Description 1',
        'Description 2',
        'Final URL'
    ]

    sa360_lookup = (
        asset_file[sa360_lookup_cols]
        .drop_duplicates(subset=['Campaign'])
        .copy()
    )

    # ===============================
    # LOAD DAILY SPEND TRACKER
    # ===============================
    daily_cols = [
        'Campaign',
        'Day',
        'Clicks',
        'Impr.',
        'CTR',
        'Avg. CPC',
        'Cost + Fees'
    ]
    daily_report = daily_spend[daily_cols].copy()

    # ===============================
    # DATA CLEANING (Using .loc to avoid warnings)
    # ===============================

    # Dates
    daily_report.loc[:, 'Day'] = pd.to_datetime(
        daily_report['Day'],
        dayfirst=True,
        errors='coerce'
    )

    # Numeric columns
    num_cols = ['Clicks', 'Impr.']
    for col in num_cols:
        daily_report.loc[:, col] = (
            pd.to_numeric(
                daily_report[col].astype(str).str.replace(',', '', regex=False),
                errors='coerce'
            )
        )

    # Money columns
    money_cols = ['Avg. CPC', 'Cost + Fees']
    for col in money_cols:
        daily_report.loc[:, col] = (
            pd.to_numeric(
                daily_report[col].astype(str).str.replace(r'[\$,]', '', regex=True),
                errors='coerce'
            )
        )

    daily_report.loc[:, 'Cost'] = (
        daily_report['Cost + Fees'] / 1.0055
    ).round(2)

    # ===============================
    # MERGE (VLOOKUP STYLE)
    # ===============================
    sem_report = daily_report.merge(
        sa360_lookup,
        on='Campaign',
        how='left'
    ).copy()

    print('✅ SEM report ready with daily metrics + creatives')

    # ===============================
    # CATEGORY ASSIGNMENT (EXCEL DRIVEN)
    # ===============================
    if 'Priority' in campaign_conditions.columns:
        campaign_conditions = campaign_conditions.sort_values('Priority')

    campaign = sem_report['Campaign'].fillna('')
    sem_report.loc[:, 'Partner'] = 'OPEX'

    for _, row in campaign_conditions.iterrows():
        target_col = 'Partner Campaign' if 'Partner Campaign' in campaign_conditions.columns else 'Partner'
        category_value = row[target_col]

        must_contain = str(row['Must_Contain']).strip() if pd.notna(row['Must_Contain']) else ''
        must_not_contain = str(row['Must_Not_Contain']).strip() if pd.notna(row['Must_Not_Contain']) else ''

        mask = (sem_report['Partner'] == 'OPEX')

        if must_contain and must_contain.lower() != 'nan':
            keywords = [k.strip() for k in must_contain.split(',') if k.strip()]
            for kw in keywords:
                mask = mask & campaign.str.contains(kw, na=False, case=False)

        if must_not_contain and must_not_contain.lower() != 'nan':
            excl_keywords = [k.strip() for k in must_not_contain.split(',') if k.strip()]
            for kw in excl_keywords:
                mask = mask & (~campaign.str.contains(kw, na=False, case=False))

        sem_report.loc[mask, 'Partner'] = category_value

    print("SEM Report Categories Assigned Successfully from Excel Rules")

    # =====================================================
    # DATE FILTER
    # =====================================================
    sem_report.loc[:, "Day"] = pd.to_datetime(
        sem_report["Day"],
        dayfirst=True,
        errors="coerce"
    )

    q_budget.loc[:, "Partner ASD"] = pd.to_datetime(
        q_budget["Partner ASD"],
        errors="coerce"
    )

    q_budget.loc[:, "Partner AED"] = pd.to_datetime(
        q_budget["Partner AED"],
        errors="coerce"
    )

    start_dates = q_budget.set_index("Partner")["Partner ASD"].to_dict()
    end_dates = q_budget.set_index("Partner")["Partner AED"].to_dict()

    sem_report.loc[:, "Start Date"] = sem_report["Partner"].map(start_dates)
    sem_report.loc[:, "End Date"] = sem_report["Partner"].map(end_dates)

    sem_report = sem_report[
        (sem_report["Day"] >= sem_report["Start Date"]) & 
        (sem_report["Day"] <= sem_report["End Date"])
    ].copy()

    sem_report.loc[:, 'Cost + Fees'] = pd.to_numeric(sem_report['Cost + Fees'], errors='coerce')

    # =====================================================
    # REPORT MONTH FILTER
    # =====================================================
    report_month = input("Enter report month (YYYY-MM), for example 2026-09: ").strip()

    try:
        report_period = pd.Period(report_month, freq="M")
    except ValueError:
        raise ValueError(
            f"Invalid month '{report_month}'. "
            "Please use YYYY-MM format, for example: 2026-09"
        )

    sem_report = sem_report[
        sem_report["Day"].dt.to_period("M") == report_period
    ].copy()

    print(f"\n📅 Generating report for: {report_period}")
    print(f"📊 Rows included: {len(sem_report):,}")

    partner_periods = (
        sem_report
        .dropna(subset=["Day"])
        .groupby("Partner")["Day"]
        .agg(
            Start_Date="min",
            End_Date="max"
        )
        .to_dict(orient="index")
    )

    # ===============================
    # AGGREGATE SEM REPORT
    # ===============================
    group_cols = [
        'Campaign', 'Partner', 'Category', 
        'Headline 1', 'Headline 2', 'Description', 'Description 1', 'Description 2', 'Final URL'
    ]
    existing_group_cols = [col for col in group_cols if col in sem_report.columns]

    def first_valid_text(series):
        valid_values = series[
            pd.notna(series) & 
            (series.astype(str).str.strip() != '') & 
            (series.astype(str).str.strip() != '--')
        ]
        if not valid_values.empty:
            return valid_values.iloc[0]
        return series.iloc[0] if not series.empty else ''

    agg_dict = {
        'Clicks': 'sum',
        'Impr.': 'sum',
        'Cost + Fees': 'sum',
        'Cost': 'sum',
        'Headline 1': first_valid_text,
        'Headline 2': first_valid_text,
        'Description': first_valid_text,
        'Description 1': first_valid_text,
        'Description 2': first_valid_text,
        'Final URL': first_valid_text
    }

    valid_agg_dict = {col: op for col, op in agg_dict.items() if col in sem_report.columns}

    sem_report_aggregated = (
        sem_report
        .groupby(existing_group_cols, as_index=False)
        .agg(valid_agg_dict)
    )

    if 'Clicks' in sem_report_aggregated.columns and 'Impr.' in sem_report_aggregated.columns:
        sem_report_aggregated.loc[:, 'CTR'] = np.where(
            sem_report_aggregated['Impr.'] > 0, 
            sem_report_aggregated['Clicks'] / sem_report_aggregated['Impr.'], 
            0
        )

    if 'Cost + Fees' in sem_report_aggregated.columns and 'Clicks' in sem_report_aggregated.columns:
        sem_report_aggregated.loc[:, 'Avg. CPC'] = np.where(
            sem_report_aggregated['Clicks'] > 0, 
            sem_report_aggregated['Cost + Fees'] / sem_report_aggregated['Clicks'], 
            0
        )

    sem_report = sem_report_aggregated.copy()
    print("SEM Report aggregated successfully, skipping '--' and blanks for text fields.")

    # ===============================
    # EXPORT CAMPAIGNS CSV
    # ===============================
    campaigns_to_highlight = sem_report[['Partner', 'Campaign']].copy()

    def export_campaigns_to_highlight(camp_df):
        base_path = os.getcwd()
        output_folder = os.path.join(base_path, 'Output')
        os.makedirs(output_folder, exist_ok=True)
        output_file = os.path.join(output_folder, 'Campaigns to Highlight.csv')
        camp_df.to_csv(output_file, index=False)
        print('Shazam! The Campaigns CSV is Ready!')

    export_campaigns_to_highlight(campaigns_to_highlight)

    # ===============================
    # CHANNEL ASSIGNMENT
    # ===============================
    c = sem_report['Campaign']
    sem_report.loc[:, 'Channel'] = np.select(
        [c.str.contains('Google', na=False), c.str.contains('Bing', na=False)],
        ['Google', 'Bing'],
        default='Null'
    )

    # ===============================
    # FINAL COLUMNS DEFINITION
    # ===============================
    final_columns = [
        'Campaign',
        'Headline 1', 'Headline 2',
        'Description', 'Description 1', 'Description 2',
        'Clicks', 'Impr.', 'CTR', 'Avg. CPC',
        'Cost', 'Cost + Fees', 'Final URL'
    ]

    reports_separated = {
        f"{cat}_{ch}": df[final_columns].copy()
        for cat, df_cat in sem_report.groupby('Partner')
        for ch, df in df_cat.groupby('Channel')
    }

    # =====================================================
    # EXCEL REPORT GENERATION
    # =====================================================
    folder_input = input("Enter report period / folder name, for example 'September FY26': ").strip()
    if not folder_input:
        folder_input = "September FY26"

    output_folder = f"{folder_input} Monthly Reports"
    os.makedirs(output_folder, exist_ok=True)

    logos_folder = "Logos"
    google_logo = os.path.join(logos_folder, "Google_Logo.png")
    bing_logo = os.path.join(logos_folder, "Bing_Logo.png")

    def get_logo_path(report_name):
        lower_name = report_name.lower()
        if 'bing' in lower_name:
            return bing_logo
        return google_logo

    for name, report in reports_separated.items():
        if report.empty:
            continue
        if "Cost + Fees" not in report.columns:
            continue

        total_cost = round(report['Cost + Fees'].sum(), 2)
        if total_cost == 0:
            continue

        total_cost_formatted = f"${total_cost:,.2f}"

        partner_lookup = (
            name
            .removesuffix("_Google")
            .removesuffix("_Bing")
            .strip()
        )

        partner_dates = partner_periods.get(partner_lookup)

        if partner_dates:
            partner_start = partner_dates["Start_Date"]
            partner_end = partner_dates["End_Date"]

            if pd.notna(partner_start) and pd.notna(partner_end):
                p_start_dt = pd.to_datetime(partner_start)
                p_end_dt = pd.to_datetime(partner_end)

                if p_start_dt.date() == p_end_dt.date():
                    period_text = p_start_dt.strftime('%B %d, %Y')
                else:
                    period_text = (
                        f"{p_start_dt.strftime('%B %d, %Y')} to "
                        f"{p_end_dt.strftime('%B %d, %Y')}"
                    )
            else:
                period_text = "Active Partner Period"
        else:
            period_text = "Reporting period unavailable"

        file_name = f"{name}_SEM_Total_Cost_{total_cost_formatted}_{folder_input}.xlsx"
        file_path = os.path.join(output_folder, file_name)

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            money_format = workbook.add_format({'num_format': '$#,##0.00'})

            main_sheet_name = "Campaign Report"

            header_df = pd.DataFrame({
                'Report Name': [f'Partner: {name}'],
                'Period': [period_text]
            })
            header_df.to_excel(writer, sheet_name=main_sheet_name, index=False, startrow=0)

            report.to_excel(writer, sheet_name=main_sheet_name, index=False, startrow=2)
            worksheet = writer.sheets[main_sheet_name]

            logo_path = get_logo_path(name)
            if os.path.exists(logo_path):
                worksheet.insert_image('E1', logo_path, {'x_scale': 0.5, 'y_scale': 0.4, 'x_offset': 100, 'y_offset': 5})

            cost_breakdown = (
                report.groupby('Campaign', as_index=False)
                .agg({
                    'Impr.': 'sum',
                    'Clicks': 'sum',
                    'Cost': 'sum',
                    'Cost + Fees': 'sum'
                })
            )

            cost_breakdown_sheet = "Cost_Breakdown"
            cost_breakdown.to_excel(
                writer,
                sheet_name=cost_breakdown_sheet,
                index=False
            )
            cost_ws = writer.sheets[cost_breakdown_sheet]

            for col in ['Cost', 'Cost + Fees', 'Avg. CPC']:
                if col in report.columns:
                    col_idx = report.columns.get_loc(col)
                    worksheet.set_column(col_idx, col_idx, None, money_format)

            cost_ws.set_column('D:E', None, money_format)

        print(f"✅ Report generated: {file_name} | Period: {period_text} | Total Cost: {total_cost_formatted}")

    print("\n✅ All reports are ready")
    print(f"📁 Output folder: {output_folder}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n[ERROR] Something went wrong:")
        traceback.print_exc()
    finally:
        input("\nClick ENTER to close this window...")

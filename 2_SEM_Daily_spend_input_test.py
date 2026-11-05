#!/usr/bin/env python
# coding: utf-8

# =====================================================
# Spend SEM
# VERSION 2
# =====================================================

import pandas as pd
import numpy as np
import os
import traceback

#---------------------------------------------------
# FUNCIÓN DE LIMPIEZA ROBUSTA PARA SÍMBOLOS Y COMAS ($ ,)
#---------------------------------------------------
def sanitize_numeric_columns(df, columns_to_check):
    
    df = df.copy()
    for col in columns_to_check:
        if col in df.columns:
            temp_series = df[col].astype(str).str.strip()
            temp_series = (
                temp_series
                .str.replace('$', '', regex=False)
                .str.replace(',', '', regex=False)
                .str.replace('€', '', regex=False)
            )
            df[col] = pd.to_numeric(temp_series, errors='coerce').fillna(0.0)
    return df

#---------------------------------------------------
# LOAD
#---------------------------------------------------

def load_inputs():
    budget_file = 'SEM Budget Input.xlsx'
    daily_spend_file = 'SEM Daily Spend.xlsx'
    q_budget = pd.read_excel(budget_file, sheet_name = 'QB', skiprows = 7)
    campaign_conditions = pd.read_excel(budget_file, sheet_name = 'Campaign_conditions')
    reference = pd.read_excel(budget_file, sheet_name='Reference table')
    daily_spend_sa360 = pd.read_excel(daily_spend_file, sheet_name='SEM Daily Spend SA360', skiprows=2)
    daily_spend_AdNet = pd.read_excel(daily_spend_file, sheet_name='Ad.net')

    return (q_budget, campaign_conditions, reference, daily_spend_sa360, daily_spend_AdNet)


#---------------------------------------------------
# PREPARE DAILY SPEND
#---------------------------------------------------

def prepare_daily_spend(daily_spend_sa360, daily_spend_AdNet, campaign_conditions):
    daily_spend_sa360 = daily_spend_sa360.copy()
    daily_spend_AdNet = daily_spend_AdNet.copy()

    #----------------------------------
    # AD.NET PREP
    #----------------------------------
    daily_spend_AdNet["Account name"] = "Ad.Net"
    rename_map = {"Date": "Day","Cost": "Cost + Fees","CPC": "Avg. CPC", "Impressions":"Impr."}
    daily_spend_AdNet = daily_spend_AdNet.rename(columns=rename_map)

    #----------------------------------
    # AD.NET CATEGORY
    #----------------------------------
    a = daily_spend_AdNet["Campaign"]

    adnet_conditions = [
        a.str.contains("Hewlett Packard BPS", na=False),
        a.str.contains("Hewlett Packard CPS Monitors", na=False),
        a.str.contains("Hewlett Packard CPS", na=False)
    ]

    adnet_choices = [
        "BPS Intel CCF Commercial",
        "CPS OPEX",
        "CPS Intel CCF Consumer"
    ]

    daily_spend_AdNet["Partner"] = np.select(adnet_conditions, adnet_choices, default="OPEX")

    #----------------------------------
    # SA360 CATEGORY (Driven by Excel Table)
    #----------------------------------
    if 'Priority' in campaign_conditions.columns:
        campaign_conditions = campaign_conditions.sort_values('Priority')

    campaign = daily_spend_sa360["Campaign"].fillna('')
    daily_spend_sa360["Partner"] = "OPEX"

    for _, row in campaign_conditions.iterrows():
        partner = row['Partner']
        must_contain = str(row['Must_Contain']).strip() if pd.notna(row['Must_Contain']) else ''
        must_not_contain = str(row['Must_Not_Contain']).strip() if pd.notna(row['Must_Not_Contain']) else ''

        mask = (daily_spend_sa360["Partner"] == "OPEX")

        if must_contain and must_contain.lower() != 'nan':
            keywords = [k.strip() for k in must_contain.split(',') if k.strip()]
            for kw in keywords:
                mask = mask & campaign.str.contains(kw, na=False, case=False)

        if must_not_contain and must_not_contain.lower() != 'nan':
            excl_keywords = [k.strip() for k in must_not_contain.split(',') if k.strip()]
            for kw in excl_keywords:
                mask = mask & (~campaign.str.contains(kw, na=False, case=False))

        daily_spend_sa360.loc[mask, "Partner"] = partner

    print("Daily Spend Prepared")
    return (daily_spend_sa360, daily_spend_AdNet)


#---------------------------------------------------
# APPEND DAILY SPEND
#---------------------------------------------------

def append_daily_spend(daily_spend_sa360, daily_spend_AdNet):
    common_columns = sorted(set(daily_spend_sa360.columns) | set(daily_spend_AdNet.columns))

    daily_spend_sa360 = daily_spend_sa360.reindex(columns=common_columns)
    daily_spend_AdNet = daily_spend_AdNet.reindex(columns=common_columns)

    daily_spend = pd.concat([daily_spend_sa360, daily_spend_AdNet], ignore_index=True, sort=False)
    print("Daily Spend Appended")
    return daily_spend


def prepare_activity_dates(q_budget):
    activity_dates = q_budget[['Partner','Partner ASD','Partner AED']].drop_duplicates()
    activity_dict = activity_dates.set_index('Partner')[['Partner ASD','Partner AED']].to_dict('index')
    print('Activity Dates Prepared')
    return activity_dict


def prepare_daily_spend_dates(daily_spend):
    date_columns = ['Day', 'Week (Sun - Sat)']
    for col in date_columns:
        if col in daily_spend.columns:
            daily_spend[col] = pd.to_datetime(daily_spend[col], errors='coerce')
    print('Daily Spend Dates Prepared')
    return daily_spend


#---------------------------------------------------
# RE-ASSIGN INACTIVE PARTNERS
#---------------------------------------------------

def re_assign_partner(daily_spend, activity_dict):
    daily_spend = daily_spend.copy()

    def assign_partner(partner, category, day):
        info = activity_dict.get(partner)
        if info is None:
            return partner

        start = info['Partner ASD']
        end = info['Partner AED']

        if pd.isna(start) or pd.isna(end):
            return partner

        if start <= day <= end:
            return partner

        if category == 'BPS':
            return 'BPS OPEX'
        if category == 'CPS':
            return 'CPS OPEX'

        return partner

    daily_spend['Partner'] = daily_spend.apply(lambda r: assign_partner(r['Partner'], r['Category'], r['Day']), axis=1)
    print('Inactive Partners Reassigned')
    return daily_spend


#---------------------------------------------------
# PARTNER ALLOCATION
#---------------------------------------------------

def allocate_partner(group, spend_col='Cost + Fees'):
    acc1 = acc2 = acc3 = 0
    out = []

    group = group.sort_values('Day')

    for _, r in group.iterrows():
        spend = r[spend_col]

        # PB1
        ps1 = 0
        if (
            pd.notna(r['PB1 Starting Date'])
            and r['PB1 Starting Date'] <= r['Day'] <= r['PB1 End Date']
            and acc1 < r['PB1']
        ):
            target = spend * r['PB1 Rate']
            ps1 = min(target, r['PB1'] - acc1)

        remaining = spend
        if r['PB1 Rate'] > 0:
            remaining -= ps1 / r['PB1 Rate']

        # PB2
        ps2 = 0
        if (
            pd.notna(r['PB2 Starting Date'])
            and r['PB2 Starting Date'] <= r['Day'] <= r['PB2 End Date']
            and acc2 < r['PB2']
        ):
            target = remaining * r['PB2 Rate']
            ps2 = min(target, r['PB2'] - acc2)

        if r['PB2 Rate'] > 0:
            remaining -= ps2 / r['PB2 Rate']

        # PB3
        ps3 = 0
        if (
            pd.notna(r['PB3 Starting Date'])
            and r['PB3 Starting Date'] <= r['Day'] <= r['PB3 End Date']
            and acc3 < r['PB3']
        ):
            target = remaining * r['PB3 Rate']
            ps3 = min(target, r['PB3'] - acc3)

        if r['PB3 Rate'] > 0:
            remaining -= ps3 / r['PB3 Rate']

        partner_spend = ps1 + ps2 + ps3

        out.append({
            **r,
            'Partner': r['Partner'],
            'PS1': ps1,
            'PS2': ps2,
            'PS3': ps3,
            'Partner Spend': partner_spend,
            'Unused Spend': remaining
        })

        acc1 += ps1
        acc2 += ps2
        acc3 += ps3

    return pd.DataFrame(out)


def allocate_partner_all(daily_spend, spend_col='Cost + Fees'):
    processed_groups = []

    for partner_name, group in daily_spend.groupby('Partner', group_keys=False):
        group = group.copy()
        group['Partner'] = partner_name
        res_group = allocate_partner(group, spend_col)
        processed_groups.append(res_group)

    daily_spend = pd.concat(processed_groups, ignore_index=True)
    print('Partner Allocation Complete')
    return daily_spend


#------------------------------------------------------
# REDIRECT WORKSTATION
#------------------------------------------------------

def redirect_workstation(daily_spend):
    daily_spend = daily_spend.copy()

    daily_spend['WS Transfer'] = np.where(
        daily_spend['Partner'] == 'BPS Intel CCF Workstation',
        daily_spend['Unused Spend'],
        0
    )

    ws_transfer = daily_spend.groupby('Day')['WS Transfer'].sum()
    daily_spend['WS Transfer Total'] = daily_spend['Day'].map(ws_transfer).fillna(0)

    daily_spend['Adjusted Spend'] = daily_spend['Cost + Fees']

    mask_ws = (daily_spend['Partner'] == 'BPS Intel CCF Workstation')
    daily_spend.loc[mask_ws, 'Adjusted Spend'] = (
        daily_spend.loc[mask_ws, 'Cost + Fees'] - daily_spend.loc[mask_ws, 'Unused Spend']
    )

    mask_comm = (daily_spend['Partner'] == 'BPS Intel CCF Commercial')
    daily_spend.loc[mask_comm, 'Adjusted Spend'] = (
        daily_spend.loc[mask_comm, 'Cost + Fees'] + daily_spend.loc[mask_comm, 'WS Transfer Total']
    )

    print('Workstation Redirect Complete')
    return daily_spend


#---------------------------------------------------------------------
# REGLA 2: REDISTRIBUCIÓN GLOBAL DE OPEX
#---------------------------------------------------------------------

def finalize_global_opex_exact_pq(daily_spend):
    df = daily_spend.copy()

    df['Final Partner Spend'] = pd.to_numeric(df['Partner Spend'], errors='coerce').fillna(0.0)
    df['Adjusted Spend'] = pd.to_numeric(df['Adjusted Spend'], errors='coerce').fillna(0.0)

    df['Final OPEX Spend'] = df['Adjusted Spend'] - df['Final Partner Spend']
    df.loc[df['Partner'] == 'BPS Intel CCF Workstation', 'Final OPEX Spend'] = 0.0

    df['Day_Str'] = df['Day'].astype(str)

    budget_cols = ['OPEX Budget', 'Aditional OPEX', 'PB1', 'PB2', 'PB3']
    for col in budget_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    def get_matrix_target(partner):
        p_str = str(partner).strip()
        if p_str in ["BPS Intel CCF Workstation", "BPS OPEX", "CPS OPEX"]:
            return None
        if "BPS" in p_str:
            return "BPS OPEX"
        elif "CPS" in p_str:
            return "CPS OPEX"
        return None

    df['OPEX Target'] = df['Partner'].apply(get_matrix_target)
    df['Budget Total Real'] = df[budget_cols].sum(axis=1)
    df = df.sort_values(['Partner', 'Day_Str']).reset_index(drop=True)
    df['AmountToTransfer'] = 0.0

    for partner, group in df.groupby('Partner'):
        if pd.isna(group['OPEX Target'].iloc[0]):
            continue

        adjusted_spend_arr = group['Adjusted Spend'].values
        budget_total_real = group['Budget Total Real'].iloc[0]

        accumulated_spend = 0.0
        excedent_amounts = []

        for current_spend in adjusted_spend_arr:
            prev_accumulated = accumulated_spend
            accumulated_spend += current_spend

            ya_estaba_agotado = prev_accumulated >= budget_total_real

            if budget_total_real == 0:
                excedent = current_spend
            elif ya_estaba_agotado:
                excedent = current_spend
            elif accumulated_spend > budget_total_real:
                excedent = accumulated_spend - budget_total_real
            else:
                excedent = 0.0

            excedent_amounts.append(excedent)

        df.loc[group.index, 'AmountToTransfer'] = excedent_amounts

    mask_transfers = (df['OPEX Target'].notna()) & (df['AmountToTransfer'] > 0)

    grouped_transfers = (
        df[mask_transfers]
        .groupby(['Day_Str', 'OPEX Target'])['AmountToTransfer']
        .sum()
        .reset_index()
        .rename(columns={'AmountToTransfer': 'Transfer OPEX', 'OPEX Target': 'Partner'})
    )

    df = df.merge(grouped_transfers, on=['Day_Str', 'Partner'], how='left')
    df['Transfer OPEX'] = df['Transfer OPEX'].fillna(0.0)

    mask_hijas = df['OPEX Target'].notna()
    df.loc[mask_hijas, 'Final OPEX Spend'] = df.loc[mask_hijas, 'Final OPEX Spend'] - df.loc[mask_hijas, 'AmountToTransfer']

    mask_matrices = df['Partner'].isin(['BPS OPEX', 'CPS OPEX'])
    df.loc[mask_matrices, 'Final OPEX Spend'] = df.loc[mask_matrices, 'Final OPEX Spend'] + df.loc[mask_matrices, 'Transfer OPEX']

    cols_to_drop = ['OPEX Target', 'AmountToTransfer', 'Transfer OPEX', 'Day_Str', 'Budget Total Real']
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])

    print("M logic successfully mirrored with strict CPS/BPS fix!")
    return df


# =====================================================
# MAIN EXECUTION BLOCKS WITH SAFE TRACEBACK & INPUT
# =====================================================

if __name__ == "__main__":
    try:
        # 1. Load Inputs
        (q_budget, campaign_conditions, reference, daily_spend_sa360, daily_spend_AdNet) = load_inputs()
        print("Inputs succesfully loaded")

        # 1.1 SANEAMIENTO INICIAL DE DATOS (Remueve $, comas y espacios que rompen cálculos)
        daily_spend_sa360 = sanitize_numeric_columns(daily_spend_sa360, ['Cost', 'Cost + Fees', 'Clicks', 'Impressions', 'Spend'])
        daily_spend_AdNet = sanitize_numeric_columns(daily_spend_AdNet, ['Cost', 'Cost + Fees', 'CPC', 'Impressions', 'Clicks'])
        q_budget = sanitize_numeric_columns(q_budget, ['Funds', 'Aditional OPEX', 'PB1', 'PB1 Rate', 'PB2', 'PB2 Rate', 'PB3', 'PB3 Rate'])

        # 2. AdNet Pre-processing & Week mapping
        daily_spend_AdNet['Account name'] = 'Ad.Net'
        rename_map = {'Date':'Day', 'Cost':'Cost + Fees','CPC':'Avg. CPC'}
        daily_spend_AdNet = daily_spend_AdNet.rename(columns=rename_map)
        daily_spend_AdNet['Week (Sun - Sat)'] = (
            daily_spend_AdNet['Day']
            - pd.to_timedelta(
                (daily_spend_AdNet['Day'].dt.dayofweek + 1) % 7,
                unit='D'
            )
        )

        # 3. Prepare & Append Daily Spend
        daily_spend_sa360, daily_spend_AdNet = prepare_daily_spend(daily_spend_sa360, daily_spend_AdNet, campaign_conditions)
        daily_spend = append_daily_spend(daily_spend_sa360, daily_spend_AdNet)

        # 4. Activity Dates & Dates Format
        activity_dict = prepare_activity_dates(q_budget)
        daily_spend = prepare_daily_spend_dates(daily_spend)

        # 5. Category Assignment
        c = daily_spend['Partner'].fillna('')
        conditions = [
            c.str.contains('CPS'),
            c.str.contains('BPS'),
            c.str.contains('PHW'),
            c.str.contains('Brand')
        ]
        choices = ['CPS', 'BPS', 'Print', 'Core']
        daily_spend['Category'] = np.select(conditions, choices, default='Uncategorized')

        # 6. Re-assign Inactive Partners
        daily_spend = re_assign_partner(daily_spend, activity_dict)

        # 7. Groupby & Budget Merge
        daily_spend = daily_spend.groupby(['Partner','Day','Week (Sun - Sat)'], as_index=False).agg({
            'HHO Revenue': 'sum',
            'Cost + Fees': 'sum',
            'Clicks': 'sum',
            'Impr.': 'sum',
            'HHO GM (Conv Time)': 'sum'
        })

        daily_spend = daily_spend.merge(q_budget[[
            'Partner', 'Category', 'Funds', 'Aditional OPEX', 'PB1', 
            'PB1 Starting Date', 'PB1 End Date', 'PB1 Rate', 'PB2',
            'PB2 Starting Date', 'PB2 End Date', 'PB2 Rate', 'PB3', 'PB3 Rate',
            'PB3 Starting Date', 'PB3 End Date'
        ]], how='left', left_on='Partner', right_on='Partner')

        daily_spend = daily_spend.rename(columns={'Funds': 'OPEX Budget'})
        daily_spend = daily_spend.sort_values(["Partner", "Day"]).reset_index(drop=True)

        # 8. Fix Types (Numérica y Fechas)
        numeric_cols = [
            'Cost + Fees', 
            'OPEX Budget', 
            'Aditional OPEX', 
            'PB1',
            'PB1 Rate', 
            'PB2',
            'PB2 Rate',
            'PB3',
            'PB3 Rate'
        ]

        for col in numeric_cols:
            if col in daily_spend.columns:
                daily_spend[col] = (
                    daily_spend[col].astype(str)
                    .str.replace(',', '', regex=False)
                    .str.replace('$', '', regex=False)
                    .pipe(pd.to_numeric, errors='coerce')
                )

        date_cols = [
            'Day',
            'PB1 Starting Date',
            'PB1 End Date',
            'PB2 Starting Date',
            'PB2 End Date',
            'PB3 Starting Date',
            'PB3 End Date'
        ]

        for col in date_cols:
            if col in daily_spend.columns:
                daily_spend[col] = pd.to_datetime(daily_spend[col], errors='coerce')

        print('Types Fixed')

        # 9. Partner Allocations & Redirects
        daily_spend = allocate_partner_all(daily_spend)
        daily_spend = redirect_workstation(daily_spend)
        daily_spend = allocate_partner_all(daily_spend, spend_col='Adjusted Spend')
        daily_spend = finalize_global_opex_exact_pq(daily_spend)

        # 10. Summary & Output Generation
        daily_summary = daily_spend[[
            'Category', 'Partner', 'Day', 'Week (Sun - Sat)', 
            'HHO Revenue', 'Cost + Fees', 'Clicks', 'Impr.', 
            'HHO GM (Conv Time)', 'PS1', 'PS2', 'PS3', 
            'Final Partner Spend', 'Final OPEX Spend'
        ]]

        daily_spend_AdNet = daily_spend_AdNet.groupby(["Partner", 'Week (Sun - Sat)'], as_index=False).agg({
            "Cost + Fees": "sum", 
            "Clicks": "sum",
            "Impr.": "sum"
        })

        os.makedirs("Output", exist_ok=True)
        daily_summary_file = os.path.join("Output", "Daily Spend Output.csv")
        adnet_file = os.path.join("Output", "AdNet_Spend.csv")

        daily_summary.to_csv(daily_summary_file, index=False)
        daily_spend_AdNet.to_csv(adnet_file, index=False)
        
        print('Daily Doc ready')
        print("¡We're good to go!")

    except Exception as e:
        print("\n[ERROR] Something went wrong:")
        traceback.print_exc()

    finally:
        input("\nClick ENTER to close this window...")

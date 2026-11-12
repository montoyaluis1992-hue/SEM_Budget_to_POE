#!/usr/bin/env python
# coding: utf-8

import os
import traceback
import numpy as np
import pandas as pd


#---------------------------------------------------
# FILE PATH
#---------------------------------------------------

def get_file(filename):
    return os.path.join('Output', filename)

#---------------------------------------------------
# LOAD
#---------------------------------------------------

def load_inputs():
    # Paths
    spend_file = get_file('Daily Spend Output.csv')

    # Validate
    for file in [spend_file]:
        if not os.path.exists(file):
            raise FileNotFoundError(f'File not found: {file}')

    # Load Daily Spend csv
    daily_spend = pd.read_csv(spend_file)

    print('Input Loaded')
    return daily_spend


# =====================================================
# MAIN PROCESS
# =====================================================

def main():
    # 1. Load inputs
    daily_spend = load_inputs()

    # 2. Get user input for month and year
    month = int(input("Enter the month number you are looking for(1-12): "))
    year = int(input("Enter the year: "))

    # 3. Filter by date
    daily_spend["Day"] = pd.to_datetime(daily_spend["Day"])

    monthly_accruals = daily_spend[
        (daily_spend["Day"].dt.month == month) &
        (daily_spend["Day"].dt.year == year)
    ].copy()

    # 4. Monthly accruals aggregation
    columns_to_sum = ['Cost + Fees', 'PS1', 'PS2', 'PS3', 'Final Partner Spend', 'Final OPEX Spend']
    valid_cols_to_agg = {col: 'sum' for col in columns_to_sum if col in monthly_accruals.columns}

    monthly_accruals = (
        monthly_accruals
        .groupby(['Partner'], as_index=False)
        .agg(valid_cols_to_agg)
    )

    # 5. Export to CSV
    monthly_accruals.to_csv('Monthly_accruals.csv', index=False)
    print("✅ Monthly accruals successfully exported to 'Monthly_accruals.csv'")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n[ERROR] Something went wrong:")
        traceback.print_exc()
    finally:
        input("\nClick ENTER to close this window...")

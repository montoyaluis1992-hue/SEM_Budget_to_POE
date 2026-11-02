#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import numpy as np
import pandas as pd

#---------------------------------------------------
# FILE PATH
#---------------------------------------------------

def get_file(filename):

    return os.path.join('Output',filename)

#---------------------------------------------------
# LOAD
#---------------------------------------------------

def load_inputs():

    # Paths
    budget_file = get_file('Original Budget.xlsx')

    spend_file = get_file('Daily Spend Output.csv')

    # Validate
    for file in [budget_file, spend_file]:

        if not os.path.exists(file):

            raise FileNotFoundError(f'File not found: {file}')

    # Load Budget sheet
    budget = pd.read_excel(budget_file, sheet_name='Budget')

    # Load Daily Spend csv
    daily_spend = pd.read_csv(spend_file)

    print('Inputs Loaded')

    return (budget, daily_spend)


# RUN
(budget,daily_spend) = load_inputs()


# In[2]:


daily_spend = daily_spend.rename(columns={'Final Partner Spend':'TPS', 'Final OPEX Spend':'OPEX Spend'})
print(daily_spend.columns)


# In[3]:


#---------------------------------------------------
# PREPARE WEEKLY SPEND
#---------------------------------------------------

def prepare_weekly_spend(budget,daily_spend):

    # Fix dates

    daily_spend['Week (Sun - Sat)'] = (pd.to_datetime(daily_spend['Week (Sun - Sat)'],errors='coerce'))

    budget['FW Start Date'] = (pd.to_datetime(budget['FW Start Date'],errors='coerce'))

    # Group

    weekly_spend = (daily_spend.groupby(['Partner',
                                         'Category',
                                         'Week (Sun - Sat)'],as_index=False).agg({'TPS':'sum',
                                                                                  'OPEX Spend':'sum'}))

    # Remove current week

    latest_week = (weekly_spend['Week (Sun - Sat)'].max())

    weekly_spend = (weekly_spend[weekly_spend['Week (Sun - Sat)']!=latest_week])

    # Merge

    budget = (budget.merge(weekly_spend[['Partner',
                                         'Week (Sun - Sat)',
                                         'TPS',
                                         'OPEX Spend']],
                           how='left',
                           left_on=['Partner','FW Start Date'],
                           right_on=['Partner','Week (Sun - Sat)'],                           
                           validate='many_to_one').drop(columns='Week (Sun - Sat)'))

    # Clean

    budget[['TPS','OPEX Spend']] = budget[['TPS','OPEX Spend']].fillna(0)

    budget = (budget.sort_values(['Partner',
                                  'Category',
                                  'FW Start Date']).reset_index(drop=True))

    print('Weekly Spend Prepared')

    return budget

budget = prepare_weekly_spend(budget,daily_spend)


# In[4]:


print(budget.columns)


# In[5]:


#---------------------------------------------------
# CALCULATE WEEKLY BUDGETS
#---------------------------------------------------

def calculate_weekly_totals(budget):

    # Total Partner Budget

    budget['Total Partner Budget'] = (budget[['WPB1',
                                              'WPB2',
                                              'WPB3']].sum(axis=1))

    # OPEX Weekly Budget

    budget['OPEX Weekly Budget'] = (budget['Total Weekly Budget']-budget['Total Partner Budget'])

    # Clean both together

    cols = ['Total Partner Budget','OPEX Weekly Budget']

    budget[cols] = (budget[cols].clip(lower=0).round(6))

    print('Weekly Budgets Calculated')

    return budget


budget = calculate_weekly_totals(budget)


# In[6]:


#---------------------------------------------------
# CALCULATE ADJUSTED DISTRIBUTION
#---------------------------------------------------

def calculate_adjusted_distribution(budget):

    #----------------------------------
    # Total Spend
    #----------------------------------

    budget['Total Spend'] = (budget[['TPS','OPEX Spend']].sum(axis=1))

    #----------------------------------
    # TWB
    #----------------------------------

    budget['TWB'] = (budget[['Total Partner Budget','OPEX Weekly Budget']].sum(axis=1))
    #----------------------------------
    # Over / Under Spend
    #----------------------------------

    budget['Partner Over/Underspend'] = ((budget['Total Partner Budget']-budget['TPS']).where(budget['TPS'] > 0,0))

    budget['OPEX Over/Underspend'] = ((budget['OPEX Weekly Budget']-budget['OPEX Spend']).where(budget
                                                                                                ['OPEX Spend'] > 0,0))

    #----------------------------------
    # Closed
    #----------------------------------

    budget['Closed'] = (budget['Total Spend']> 0)

    #----------------------------------
    # Remaining Weeks
    #----------------------------------

    partner_remaining = (budget['Total Partner Budget'].where(~budget['Closed'],0).groupby
                         (budget['Partner']).transform('sum'))

    budget['Partner Remaining Weeks'] = (partner_remaining.where(~budget['Closed'], 0))

    opex_remaining = (budget['OPEX Weekly Budget'].where(~budget['Closed'],0).groupby(budget['Partner']).transform('sum'))

    budget['OB Remaining Weeks'] = (opex_remaining.where(~budget['Closed'],0))

    #----------------------------------
    # New %
    #----------------------------------

    budget['New PB%'] = (budget['Total Partner Budget'].div(budget['Partner Remaining Weeks']).replace
                         ([np.inf,-np.inf],0).fillna(0))

    budget['New OPEX%'] = (budget['OPEX Weekly Budget'].div(budget['OB Remaining Weeks']).replace([np.inf,-np.inf],0).fillna
                           (0))

    print('Adjusted Distribution Calculated')

    return budget


budget = calculate_adjusted_distribution(
    budget
)


# In[7]:


#-----------------------------------------------------------
# CALCULATE ADJUSTED WEEKLY BUDGET (FAST)
#-----------------------------------------------------------

def calculate_adjusted_WB_fast(budget):

    #----------------------------------
    # Accumulated Over / Under Spend
    #----------------------------------

    budget[['Partner Acc. Over/Underspend',
            'OPEX Acc. Over/Underspend']] = (budget.groupby('Partner')[['Partner Over/Underspend',
                                                                        'OPEX Over/Underspend']].transform('sum'))
    #----------------------------------
    # New Totals
    #----------------------------------

    open_weeks = (budget['TPS']==0)
    budget['New PB Total'] = ((budget['Partner Remaining Weeks']+budget['Partner Acc. Over/Underspend']).where(open_weeks,0))

    budget['New Opex Total'] = ((budget['OB Remaining Weeks']+budget['OPEX Acc. Over/Underspend']).where(open_weeks,0))

    #----------------------------------
    # Adjusted Budget
    #----------------------------------

    budget['AdjPB'] = (budget['New PB Total']*budget['New PB%'])

    budget['AdjOB'] = (budget['New Opex Total']*budget['New OPEX%'])

    #----------------------------------
    # Final Weekly Budget
    #----------------------------------

    budget['Adjusted Total Weekly Budget'] = ((budget['AdjOB']+
                                               budget['AdjPB']).where(budget['Total Spend']==0,budget['Total Spend']))

    print('Adjusted Weekly Budget Calculated')

    return budget


budget = calculate_adjusted_WB_fast(budget)


# In[8]:


#---------------------------------------------------
# ADJUSTED WATERFALL (FAST - SAME OUTPUT)
#---------------------------------------------------

def run_adjusted_waterfall_fast(budget):

    #----------------------------------
    # Sort
    #----------------------------------

    budget = (budget.sort_values(['Partner','FW Start Date']).reset_index(drop=True))

    groups = (budget.groupby('Partner',sort=False))

    #----------------------------------
    # APB1
    #----------------------------------

    pb1_active = ((budget['FW Start Date'] >= budget['PB1 Starting Date'])&
                  (budget['FW Start Date'] <= budget['PB1 End Date']))

    budget['Adjusted Partner Spend'] = np.where(budget['Total Spend'] > 0,budget['TPS'],
                                                np.where(pb1_active,
                                                         (budget['Adjusted Total Weekly Budget']*budget['PB1 Rate']),0))

    acc1 = (groups['Adjusted Partner Spend'].cumsum()-budget['Adjusted Partner Spend'])

    budget['WAPB1'] = np.where(acc1 <= (budget['PB1']-budget['Adjusted Partner Spend']),
                               budget['Adjusted Partner Spend'],np.where(acc1 > budget['PB1'], 0,budget['PB1'] - acc1))

    #----------------------------------
    # APB2
    #----------------------------------

    acc_w1 = (groups['WAPB1'].cumsum()-budget['WAPB1'])

    pb1_spend = (budget['WAPB1'].div(budget['PB1 Rate']).fillna(0))

    base_pb2 = (budget['Adjusted Total Weekly Budget']-pb1_spend)

    budget['APB2'] = np.where(acc_w1 < (budget['PB1']-budget['WAPB1']),0,
                              np.where((budget['PB2']>0)&(budget['FW Start Date']>=budget['PB2 Starting Date']),
                                       (base_pb2*budget['PB2 Rate']),0))

    acc2 = (groups['APB2'].cumsum()-budget['APB2'])

    budget['WAPB2'] = np.where(acc2 <= (budget['PB2']-budget['APB2']),budget['APB2'],
                               np.where(acc2 > budget['PB2'],0,budget['PB2'] - acc2))

    #----------------------------------
    # APB3
    #----------------------------------

    pb2_spend = (budget['WAPB2'].div(budget['PB2 Rate']).fillna(0))

    base_pb3 = (budget['Adjusted Total Weekly Budget']-pb1_spend-pb2_spend)

    budget['APB3'] = (base_pb3*budget['PB3 Rate']).clip(lower=0)

    acc3 = (groups['APB3'].cumsum()-budget['APB3'])

    budget['WAPB3'] = np.where(acc3 <= (budget['PB3']-budget['APB3']),
                               budget['APB3'],
                               np.where(acc3 > budget['PB3'],0,budget['PB3'] - acc3))

    #----------------------------------
    # Final
    #----------------------------------

    budget['Adj TPB'] = (budget[['WAPB1',
                                 'WAPB2',
                                 'WAPB3']].sum(axis=1))

    budget['Adj OPEXWB'] = (budget['Adjusted Total Weekly Budget'] - budget['Adj TPB'])

    print('Adjusted Waterfall Complete')

    return budget


budget = run_adjusted_waterfall_fast(budget)


# In[9]:


def create_totals(budget): 
    budget['Adj TPB'] = budget['WAPB1']+budget['WAPB2']+budget['WAPB3']
    budget['Adj OPEXWB'] = budget['Adjusted Total Weekly Budget'] - budget['Adj TPB'] 
    budget['Total Partner'] = budget['PB1'] + budget['PB2'] + budget['PB3']
    budget['Total OPEX'] = budget['Funds'] + budget['Aditional OPEX']
    print("Totals ready")
    return budget


budget = create_totals(budget)


# In[10]:


budget_final_doc = (budget[['Partner', 'Category', 'Fiscal Week', 'FW Start Date', 'Funds', 'Aditional OPEX', 
                           'PB1', 'PB2', 'PB3','Total Partner', 'Total OPEX','WPB1', 'WPB2', 
                            'WPB3', 'Total Partner Budget','OPEX Weekly Budget', 'TPS', 'OPEX Spend', 
                            'WAPB1',  'WAPB2',  'WAPB3','Adj TPB','Adj OPEXWB']]).copy()


# In[11]:


##ORIGINAL BUDGET##
#---------------------------------
# PARTNER
#---------------------------------

partner = (budget_final_doc[["Partner", 'Category', "Fiscal Week", 'Total Partner Budget']].copy())
partner["Funds"] = "Partner"
partner = partner.rename(columns={"Total Partner Budget": "Amount"})

#---------------------------------
# OPEX
#---------------------------------

opex = budget_final_doc[['Partner', 'Category', 'Fiscal Week', 'OPEX Weekly Budget']].copy()
opex["Funds"] = "OPEX"
opex = opex.rename(columns={"OPEX Weekly Budget": "Amount"})

#---------------------------------
# Concat
#---------------------------------
original_budget = pd.concat([partner, opex], ignore_index=True)

excluded_categories = [
    'Print',
    'PLA PSG BPS',
    'PLA PSG CPS'
]

original_budget = original_budget[

    ~(
        (original_budget['Funds'] == 'Partner')
        &
        (
            original_budget['Category']
            .isin(
                excluded_categories
            )
        )

    )

].copy()

#---------------------------------
# PIVOT
#---------------------------------

original_budget_pivot = (original_budget.pivot_table(index=["Partner", 'Category', "Funds"],
                                                     columns="Fiscal Week",
                                                     values="Amount",
                                                     aggfunc="sum",
                                                     fill_value=0).reset_index())

original_budget_pivot.columns.name = None

original_budget_pivot = (original_budget_pivot.sort_values(['Category', "Partner", "Funds"])).reset_index(drop=True)


# In[12]:


##Adjusted BUDGET##
# PARTNER
#---------------------------------
partner_adj = budget_final_doc[['Partner', 'Category', 'Fiscal Week', 'Adj TPB']].copy()
partner_adj['Funds'] = 'Partner'
partner_adj = partner_adj.rename(columns={'Adj TPB': 'Amount'})


#OPEX Adjusted

opex_adj = budget_final_doc[['Partner', 'Category', 'Fiscal Week', 'Adj OPEXWB']].copy()
opex_adj['Funds'] = 'OPEX'
opex_adj = opex_adj.rename(columns={'Adj OPEXWB': 'Amount'})

### Concat

adjusted_budget = pd.concat([partner_adj, opex_adj], ignore_index=True)

#Clean OPEX ONLY campaigns
excluded_categories = [
    'Print',
    'PLA PSG BPS',
    'PLA PSG CPS'
]

adjusted_budget = adjusted_budget[

    ~(
        (adjusted_budget['Funds'] == 'Partner')
        &
        (
            adjusted_budget['Category']
            .isin(
                excluded_categories
            )
        )

    )

].copy()

##Pivot

adjusted_budget_pivot = (adjusted_budget.pivot_table(index=['Partner', 'Category', 'Funds'], 
                                                     columns='Fiscal Week', 
                                                     values='Amount', 
                                                     aggfunc='sum', 
                                                     fill_value=0).reset_index())


adjusted_budget_pivot.columns.name = None
adjusted_budget_pivot = (adjusted_budget_pivot.sort_values(['Category', 'Partner','Funds'])).reset_index(drop=True)


# In[13]:


import os
import pandas as pd

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


#---------------------------------------------------
# FILE PATH
#---------------------------------------------------

def get_file(
    filename
):

    output_folder = 'Output'

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    return os.path.join(
        output_folder,
        filename
    )


#---------------------------------------------------
# EXPORT OUTPUT
#---------------------------------------------------

output_path = get_file(
    'Budget_Output2.0.xlsx'
)

with pd.ExcelWriter(

    output_path,

    engine='openpyxl'

) as writer:

    sheets = {

        'Original Budget': original_budget_pivot,

        'Adjusted Budget': adjusted_budget_pivot,

        'All Data': budget_final_doc

    }

    for sheet_name, df in sheets.items():

        # Export dataframe

        df.to_excel(

            writer,

            sheet_name=sheet_name,

            index=False

        )

        ws = writer.sheets[sheet_name]

        #----------------------------------
        # Excel Table
        #----------------------------------

        max_row = len(df) + 1

        max_col = len(df.columns)

        table_range = (

            f"A1:"

            f"{get_column_letter(max_col)}"

            f"{max_row}"

        )

        table_name = (

            sheet_name

            .replace(
                " ",
                "_"
            )

            .replace(
                "-",
                "_"
            )

        )

        tab = Table(

            displayName=table_name,

            ref=table_range

        )

        style = TableStyleInfo(

            name="TableStyleMedium2",

            showFirstColumn=False,

            showLastColumn=False,

            showRowStripes=True,

            showColumnStripes=False

        )

        tab.tableStyleInfo = style

        ws.add_table(
            tab
        )

        #----------------------------------
        # Auto Width
        #----------------------------------

        for column in ws.columns:

            width = max(

                len(
                    str(cell.value)
                )

                if cell.value is not None

                else 0

                for cell in column

            )

            ws.column_dimensions[
                column[0].column_letter
            ].width = width + 2




print('Shazam!! Your budget is ready ✨')



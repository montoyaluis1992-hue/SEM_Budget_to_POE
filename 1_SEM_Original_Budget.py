#!/usr/bin/env python
# coding: utf-8

# In[1]:


  ######    ########    ######    ######    ######  ##      ##    ######    ##                ########    
  ######    # ######    ##  ##    ######    ##  ##  ##      ##    ######    ##                # ######    
##      ##  ##      ##    ##    ##            ##    # ##    ##  ##      ##  ##                ##      ##  
##      ##  ##      ##    ##    ##            ##    # ##    ##  ##      ##  ##                ##      ##  
##      ##  # ######      ##    ##    ####    ##    ##  ##  ##  # ###### #  ##                # ######    
##      ##  # #### #      ##    ##    ## #    ##    ##  ##  ##  # ###### #  ##                # ######    
##      ##  ##    ##      ##    ##      ##    ##    ##    ## #  ##      ##  ##                ##      ##  
##      ##  ##    ##      ##    ##      ##    ##    ##    ## #  ##      ##  ##                ##      ##  
  ######    ##      ##  ##  ##    ######    ##  ##  ##      ##  ##      ##  # ########        # ######    
  ######    ##      ##  ######    ######    ######  ##      ##  ##      ##  ##########        ########    


# In[2]:


# =====================================================
# BUDGET GENERATOR
# VERSION 2
# =====================================================

import pandas as pd
import numpy as np
import os
#---------------------------------------------------
# LOAD
#---------------------------------------------------

def load_inputs():

    budget_file = 'SEM Budget Input.xlsx'
    daily_spend_file = 'SEM Daily Spend .xlsx'

    q_budget = pd.read_excel(budget_file, sheet_name = 'QB', skiprows = 7)

    weekly_weight = pd.read_excel(budget_file, sheet_name='Weekly_Weight')

    reference = pd.read_excel(budget_file, sheet_name='Reference table')   


    return (q_budget, weekly_weight, reference)

(q_budget, weekly_weight, reference) = load_inputs()

#---------------------------------------------------
# ORIGINAL BUDGET
#---------------------------------------------------

def build_original_budget(q_budget, weekly_weight):

    budget = q_budget.copy()
    weeks = weekly_weight['Fiscal Week'].unique()

    budget = pd.DataFrame([{**row.to_dict(),'Fiscal Week': fw}
                           for _, row in q_budget.iterrows()
                           for fw in weeks])




    print('Original Budget Prepared')


    return budget

budget = build_original_budget(q_budget, weekly_weight)



# In[ ]:





# In[4]:


#---------------------------------------------------
# PREPARE DATES
#---------------------------------------------------

def prepare_dates(budget, reference):

    date_columns = [
        'Starting Date','End Date',
        'AO Starting Date','AO End Date',
        'PB1 Starting Date','PB1 End Date',
        'PB2 Starting Date','PB2 End Date',
        'PB3 Starting Date','PB3 End Date']

    # Existing columns only
    valid_cols = [col for col in date_columns if col in budget.columns]

    # Convert to datetime
    budget[valid_cols] = (budget[valid_cols].apply(pd.to_datetime,errors='coerce'))

    # Move dates to Sunday
    for col in valid_cols:
        budget[col] = (budget[col]-pd.to_timedelta((budget[col].dt.dayofweek + 1) % 7,unit='D'))

    # Fiscal Week → FW Start Date

    budget = (budget.drop(columns=['FW Start Date'], errors='ignore').merge(
        reference[['Fiscal Week','FW']],how='left',on='Fiscal Week').rename(columns={'FW':'FW Start Date'}))

    budget['FW Start Date'] = (pd.to_datetime(budget['FW Start Date'],errors='coerce'))

    print('Dates Prepared')

    return budget

budget = prepare_dates(budget, reference)


# In[5]:


#Build weekly weight table
def build_weight_table(weekly_weight, reference): 
        resultados = []
        for category in weekly_weight["Category"].unique():
                df_cat = weekly_weight[weekly_weight["Category"] == category].copy()
                weeks = df_cat["Week Number"].tolist()
                combinations = [(start, end) for start in weeks for end in weeks if start <= end]
                for start, end in combinations:
                    sub = df_cat[(df_cat["Week Number"] >= start) & (df_cat["Week Number"] <= end)].copy()
                    total = sub["Original percentage"].sum()
                    if total > 0:
                        sub["Adjusted %"] = sub["Original percentage"] / total
                        sub["Start Week"] = start
                        sub["End Week"] = end
                        sub["Category"] = category
                        resultados.append(sub)
    # Unir todos los resultados
        weekly_weight_adjusted = pd.concat(resultados, ignore_index=True)     


        #-----------------------------------------
        # Convert Week Number → Dates#
        #-----------------------------------------
        weekly_weight_adjusted = (weekly_weight_adjusted.merge(reference[['Week Number', 'FW']],
                                                               how='left',left_on='Start Week',right_on='Week Number'))

        weekly_weight_adjusted = (weekly_weight_adjusted.rename(columns={'FW': 'Start FW'}))
        weekly_weight_adjusted = (weekly_weight_adjusted.drop(columns=['Week Number'], errors='ignore'))
        weekly_weight_adjusted = (weekly_weight_adjusted.merge(reference[['Week Number','FW']], how='left',left_on='End Week',
                                                               right_on='Week Number' ))
        weekly_weight_adjusted = (weekly_weight_adjusted.rename(columns={'FW': 'End FW'}))
        weekly_weight_adjusted = (weekly_weight_adjusted.drop(columns=['Week Number_x', 'Week Number_y'], errors='ignore'))




        print('Weights Prepared')
        return (weekly_weight_adjusted)

weekly_weight_adjusted = (build_weight_table(weekly_weight, reference))


# In[6]:


########### V2
def prepare_weights_fast(budget, weekly_weight):

    weight_columns = [
        ('Starting Date','End Date','OPEX Weight'),
        ('AO Starting Date','AO End Date','ADOPEX Weight'),
        ('PB1 Starting Date','PB1 End Date','PB1 Weight'),
        ('PB2 Starting Date','PB2 End Date','PB2 Weight'),
        ('PB3 Starting Date','PB3 End Date','PB3 Weight')]

    # Crear lookup una sola vez
    adjusted_lookup = (weekly_weight_adjusted.set_index(['Category',
                                                'Fiscal Week',
                                                'Start FW',
                                                'End FW'])['Adjusted %'])

    original_lookup = (weekly_weight_adjusted.set_index(['Category',
                                                'Fiscal Week',
                                                'Start FW','End FW'])['Original percentage'])

    for start, end, output in weight_columns:

        key = pd.MultiIndex.from_frame(budget[['Category',
                                               'Fiscal Week',
                                               start,
                                               end]])
        budget[output] = (key.map(adjusted_lookup).fillna(0))

    # Original Weight

    key = pd.MultiIndex.from_frame(budget[['Category',
                                           'Fiscal Week',
                                           'Starting Date',
                                           'End Date']])

    budget['Original Weight'] = (key.map(original_lookup).fillna(0))

    print('Weights Added')

    return budget

budget = prepare_weights_fast(budget, weekly_weight_adjusted)


# In[7]:


#---------------------------------------------------
# PREPARE PARTNER WEIGHT ADJUSTED
#---------------------------------------------------

def prepare_partner_weight_adjusted(budget):

    budget = budget.copy()

    #----------------------------------
    # ACTIVE WEEKS
    #----------------------------------

    for pb in ['PB1','PB2','PB3']:

        budget[f'AW {pb}'] = np.where((budget[pb] * budget[f'{pb} Weight']) > 0,'Active Partner','No Active')

    #----------------------------------
    # OVERLAP
    #----------------------------------

    budget['Overlap weeks'] = ((budget['AW PB1'] == 'Active Partner').astype(int)+
                               (budget['AW PB2'] == 'Active Partner').astype(int)+
                               (budget['AW PB3'] == 'Active Partner').astype(int))

    #----------------------------------
    # ADJUSTMENT FACTOR
    #----------------------------------

    budget['Adjustment Factor'] = np.where(budget['Overlap weeks'] > 1,1 / budget['Overlap weeks'],1)

    #----------------------------------
    # NO OVERLAP WEEKS
    #----------------------------------
    # Excel:
    # COUNTIFS(
    # Partner,
    # AW PB1="Active Partner",
    # Overlap=1
    # )

    mask = ((budget['AW PB1'] == 'Active Partner')&(budget['Overlap weeks'] == 1))

    no_overlap = (budget.loc[mask].groupby('Partner')['Fiscal Week'].nunique())

    budget['No Overlap Weeks'] = (budget['Partner'].map(no_overlap).fillna(0))

    #----------------------------------
    # ADJ PBW
    #----------------------------------

    for pb in ['PB1','PB2','PB3']:

        weight = f'{pb} Weight'

        output = f'ADJ {pb}W'

        budget[output] = np.where(budget['No Overlap Weeks'] == 0,
                                  np.where(budget[pb] > 0,budget[weight],0),
                                  np.where(budget[pb] > 0,np.where(budget['Overlap weeks'] > 1,
                                                                   budget[weight]/budget['Overlap weeks'],budget[weight]*
                                                                   budget['Adjustment Factor']),0))

    #----------------------------------
    # ADJ %
    #----------------------------------

    for pb in ['PB1','PB2','PB3']:

        budget[f'{pb} Adj%'] = (budget[f'{pb} Weight']-budget[f'ADJ {pb}W']).clip(lower=0)

    #----------------------------------
    # REDISTRIBUTE
    #----------------------------------

    for pb in ['PB1','PB2','PB3']:

        redistribution = (budget.groupby('Partner')[f'{pb} Adj%'].transform('sum'))

        eligible = ((budget['Overlap weeks'] == 1)&(budget[f'{pb} Weight'] > 0)
                    &(budget['Adjustment Factor'] == 1))

        divisor = (eligible.groupby(budget['Partner']).transform('sum').replace(0,np.nan))

        budget[f'{pb}W ADJ'] = (redistribution.div(divisor).fillna(0))

    #----------------------------------
    # FINAL WEIGHT
    #----------------------------------

    for pb in ['PB1','PB2','PB3']:

        budget[f'{pb} WA'] = np.where((budget[f'ADJ {pb}W'] > 0)&
                                      (budget[f'{pb} Adj%'] == 0),budget[f'ADJ {pb}W']+
                                      budget[f'{pb}W ADJ'],budget[f'ADJ {pb}W'])

        budget[f'{pb} WA'] = (budget[f'{pb} WA'].fillna(0))

    print('Partner Weight Adjusted')

    return budget


budget = prepare_partner_weight_adjusted(budget)


# In[ ]:





# In[9]:


#---------------------------------------------------
# CALCULATE WEEKLY TOTAL BUDGET
#---------------------------------------------------

def calculate_weekly_total_budget(budget):

    #----------------------------------
    # MSP
    #----------------------------------

    budget['MSP'] = (budget['PB1'].div(budget['PB1 Rate']).where(budget['PB1'] > 0, 0) +
                     budget['PB2'].div(budget['PB2 Rate']).where(budget['PB2'] > 0, 0) +
                     budget['PB3'].div(budget['PB3 Rate']).where(budget['PB3'] > 0, 0)).fillna(0)

    #----------------------------------
    # Minimum Match Spend
    #----------------------------------

    pb1_mms = ((budget['PB1'] * budget['PB1 WA']).div(budget['PB1 Rate']).replace([np.inf, -np.inf], 0).fillna(0))

    pb2_mms = ((budget['PB2'] * budget['PB2 WA']).div(budget['PB2 Rate']).replace([np.inf, -np.inf], 0).fillna(0))

    pb3_mms = ( (budget['PB3'] * budget['PB3 WA']).div(budget['PB3 Rate']).replace([np.inf, -np.inf], 0).fillna(0))

    budget['Minim Match Spend'] = pb1_mms + pb2_mms + pb3_mms

    #----------------------------------
    # Match Spend Partner
    #----------------------------------

    budget['MT Spend PB1'] = (budget['PB1'] * budget['PB1 WA']).fillna(0)
    budget['MT Spend PB2'] = (budget['PB2'] * budget['PB2 WA']).fillna(0)
    budget['MT Spend PB3'] = (budget['PB3'] * budget['PB3 WA']).fillna(0)

    #----------------------------------
    # Total Partner
    #----------------------------------

    budget['MTPB'] = budget[['MT Spend PB1', 'MT Spend PB2', 'MT Spend PB3']].sum(axis=1)

    #----------------------------------
    # Additional OPEX
    #----------------------------------

    budget['AOB'] = (budget['Aditional OPEX'] *budget['ADOPEX Weight']).fillna(0)

    #----------------------------------
    # Match Spend OPEX
    #----------------------------------

    budget['MT Spend OPEX'] = (budget['Funds'] * budget['OPEX Weight']+ budget['AOB']).fillna(0)

    #----------------------------------
    # ORIGINAL WEEKLY TOTAL BUDGET
    #----------------------------------

    budget['WTB'] = (budget['MTPB'] + budget['MT Spend OPEX'])

    budget['WTB'] = budget['WTB'].clip(lower=0)

    #=========================================================
    # NEW LOGIC
    #=========================================================

    #----------------------------------
    # Expected Total Budget
    #----------------------------------

    budget['Expected TB'] = np.maximum(budget['WTB'], budget['Minim Match Spend'])

    #----------------------------------
    # Extra OPEX Needed This Week
    #----------------------------------

    budget['Over/Under OPEX Needed'] = budget['Expected TB'] - budget['WTB']

    #----------------------------------
    # Total OPEX Reallocation by Partner
    #----------------------------------

    budget['Aditional OPEX Needed'] = (budget.groupby(['Partner', 'Category'])['Over/Under OPEX Needed'].transform('sum'))

    #----------------------------------
    # OPEX Weight Available
    #----------------------------------

    available_weight = (budget.loc[budget['Over/Under OPEX Needed'] == 0].groupby(['Partner', 'Category'])
                        ['OPEX Weight'].transform('sum'))

    budget['Adjustment'] = np.where(budget['Over/Under OPEX Needed'] == 0,budget['OPEX Weight'] / available_weight,0)

    budget['Adjustment'] = (budget['Adjustment'].replace([np.inf, -np.inf], 0).fillna(0))

    #----------------------------------
    # FINAL WEEKLY TOTAL BUDGET
    #----------------------------------

    budget['WTB'] = (budget['Expected TB'] - (budget['Aditional OPEX Needed'] * budget['Adjustment']))

    budget['WTB'] = budget['WTB'].clip(lower=0)

    print('Weekly Total Budget Calculated')

    return budget


budget = calculate_weekly_total_budget(budget)


# In[ ]:





# In[10]:


#---------------------------------------------------
# RUN PARTNER WATERFALL
#---------------------------------------------------

def run_partner_waterfall(budget):

    budget = (budget.sort_values(['Partner','FW Start Date']).reset_index(drop=True))

    #----------------------------------
    # Helper
    #----------------------------------

    def allocate_layer(source,pb,rate,output):

        acc = (budget.groupby('Partner')[source].cumsum()-budget[source])

        result = np.where(acc <= (budget[pb]-budget[source]),budget[source],
                          np.where(acc > budget[pb],0,budget[pb]-acc))

        budget[output] = (pd.Series(result,index=budget.index).clip(lower=0).fillna(0))

    #----------------------------------
    # PPB
    #----------------------------------

    budget['PPB'] = np.where((budget['FW Start Date']>=budget['PB1 Starting Date'])&
                             (budget['FW Start Date']<=budget['PB1 End Date']),budget['WTB']*budget['PB1 Rate'],0)

    allocate_layer('PPB','PB1','PB1 Rate','WPB1')

    #----------------------------------
    # PPB2
    #----------------------------------

    budget['Accumulated PPWB'] = (budget.groupby('Partner')['WPB1'].cumsum()-budget['WPB1'])

    remaining = (budget['WTB']-(budget['WPB1'].div(budget['PB1 Rate'])).fillna(0))

    budget['Secondary Partner'] = np.where(budget['Accumulated PPWB']<(budget['PB1']-budget['WPB1']),0,
                                           np.where((budget['PB2']>0)&((budget['FW Start Date']>=budget['PB2 Starting Date'])&
                                                                      (budget['FW Start Date']<=budget['PB2 End Date'])),
                                                    remaining*budget['PB2 Rate'], 0))

    allocate_layer('Secondary Partner','PB2','PB2 Rate','WPB2')

    #----------------------------------
    # PPB3
    #----------------------------------

    remaining = (remaining-(budget['WPB2'].div(budget['PB2 Rate'])).fillna(0))

    budget['WB3'] = np.where((budget['PB3']>0)&
                             ((budget['FW Start Date']>=budget['PB3 Starting Date'])&
                               (budget['FW Start Date']<=budget['PB2 End Date'])),remaining*budget['PB3 Rate'],0)

    allocate_layer('WB3','PB3','PB3 Rate','WPB3')

    #----------------------------------
    # Totals
    #----------------------------------

    budget['Total Partner Budget'] = (budget[['WPB1','WPB2','WPB3']].sum(axis=1))

    print('Partner Waterfall Complete')

    return budget

budget = (run_partner_waterfall(budget))


# In[11]:


#---------------------------------------------------
# FINALIZE TOTALS
#---------------------------------------------------

def finalize_budget_totals(budget):

    budget = budget.copy()

    #----------------------------------
    # Partner Budget
    #----------------------------------

    budget['Total Partner Budget'] = (budget[['WPB1','WPB2','WPB3']].sum(axis=1).clip(lower=0))

    #----------------------------------
    # Default OPEX
    #----------------------------------

    budget['OPEX Weekly Budget'] = (budget['WTB'] - budget['Total Partner Budget'])

    budget['OPEX Weekly Budget'] = (budget['OPEX Weekly Budget'].clip(lower=0))

    #----------------------------------
    # Apply OPEX recalculated
    #----------------------------------

    mask = (budget['Partner'].isin(['BPC OPEX','CPC OPEX']))

    budget.loc[mask,'OPEX Weekly Budget'] = (budget.loc[mask,'MT Spend OPEX'])

    #----------------------------------
    # Final Total
    #----------------------------------

    budget['Total Weekly Budget'] = (budget['Total Partner Budget']+budget['OPEX Weekly Budget']).round(6)

    print('Final Totals Calculated')

    return budget


budget = (finalize_budget_totals(budget))



# In[12]:


##ORIGINAL BUDGET##
#---------------------------------
# PARTNER
#---------------------------------

partner = (budget[["Partner", 'Category', "Fiscal Week", 'Total Partner Budget']].copy())
partner["Funds"] = "Partner"
partner = partner.rename(columns={"Total Partner Budget": "Amount"})

#---------------------------------
# OPEX
#---------------------------------

opex = budget[['Partner', 'Category', 'Fiscal Week', 'OPEX Weekly Budget']].copy()
opex["Funds"] = "OPEX"
opex = opex.rename(columns={"OPEX Weekly Budget": "Amount"})

#---------------------------------
# Concat
#---------------------------------
original_budget = pd.concat([partner, opex], ignore_index=True)

excluded_categories = [
    'PHW OPEX',
    'BPS OPEX',
    'CPS OPEX', 
    'Brand OPEX'
]

original_budget = original_budget[~((original_budget['Funds'] == 'Partner')&
                                    (original_budget['Partner'].isin(excluded_categories)))].copy()

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



# In[13]:


import os

#---------------------------------------------------
# EXPORT OUTPUT
#---------------------------------------------------

def export_output(original_budget_pivot,budget):

    # Get script location
    base_path = os.getcwd()

    # Create Output folder
    output_folder = os.path.join(base_path,'Output')

    os.makedirs(output_folder,exist_ok=True)

    # Full file path
    output_file = os.path.join(output_folder,'Original Budget.xlsx')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        original_budget_pivot.to_excel(writer,sheet_name='Original Budget', index=False)

        budget.to_excel( writer,sheet_name='Budget',index=False)

    print('Shazam! The Original Budget is Ready!')


# RUN
export_output(original_budget_pivot,budget)


# In[1]:


import traceback

if __name__ == "__main__":
    try:
        # Aquí iría todo tu flujo principal de ejecución 
        # (desde la carga de inputs hasta guardar el CSV)
        print("¡We're Good to Go!")

    except Exception as e:
        print("\n[ERROR] Something went wrong:")
        traceback.print_exc()

    finally:
        # Esto se ejecuta siempre, tanto si hubo error como si terminó bien
        input("\nClick ENTER to close this window...")


# In[14]:


  ########  ########    ##########  ##      ##  ########    
  ########  # ######    # ########  ##      ##  # ######    
##          ##      ##  ##          # ##    ##  ##      ##  
##          ##      ##  ##          # ##    ##  ##      ##  
  ######    # ######    # ######    ##  ##  ##  ##      ##  
  ######    # ######    # ######    ##  ##  ##  ##      ##  
        ##  ##          ##          ##    ## #  ##      ##  
        ##  ##          ##          ##    ## #  ##      ##  
########    ##          # ########  ##      ##  # ######    
########    ##          ##########  ##      ##  ########   


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

    spend_file = get_file('Daily Spend Output.csv')

    # Validate
    for file in [spend_file]:

        if not os.path.exists(file):

            raise FileNotFoundError(f'File not found: {file}')


    # Load Daily Spend csv
    daily_spend = pd.read_csv(spend_file)

    print('Input Loaded')

    return (daily_spend)


# RUN
(daily_spend) = load_inputs()


# In[2]:


month = int(input("Enter the month number you are looking for(1-12)"))
year = int(input("Enter the year"))

daily_spend["Day"] = pd.to_datetime(daily_spend["Day"])

monthly_accruals = daily_spend[
    (daily_spend["Day"].dt.month == month) &
    (daily_spend["Day"].dt.year == year)
].copy()


# In[ ]:





# In[5]:


#---------------------------------------------------
# MONTHLY ACCRUALS AGGREGATION
#---------------------------------------------------

# Definir las columnas que queremos sumar si existen
columns_to_sum = ['Cost + Fees', 'PS1', 'PS2', 'PS3', 'Final Partner Spend', 'Final OPEX Spend']

# Filtrar solo las columnas que realmente están presentes en el DataFrame
valid_cols_to_agg = {col: 'sum' for col in columns_to_sum if col in monthly_accruals.columns}

# Realizar la agregación de forma segura
monthly_accruals = (
    monthly_accruals
    .groupby(['Partner'], as_index=False)
    .agg(valid_cols_to_agg)
)


# In[6]:


monthly_accruals.to_csv('Monthly_accruals.csv', index=False)


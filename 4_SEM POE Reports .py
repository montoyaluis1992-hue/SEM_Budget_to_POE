#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import os
import xlsxwriter


# In[2]:


################################################
###########     Get Copy & assets    ###########
################################################

def get_file(filename):

    return os.path.join('POE Report Assets',filename)

def load_inputs():

    # Paths
    asset_file = get_file('SA360_SEM_Report.csv')


    # Validate
    for file in [asset_file]:

        if not os.path.exists(file):

            raise FileNotFoundError(f'File not found: {file}')


    # Load Item sheet
    asset_file = pd.read_csv(asset_file, skiprows=2)


    print('Inputs Loaded')

    return (asset_file)
(asset_file) = load_inputs()


# In[3]:


################################################
###########     Get Spend Data       ###########
################################################

def load_inputs2():

    budget_file = 'SEM Budget Input.xlsx'
    daily_spend_file = 'SEM Daily Spend.xlsx'

    q_budget = pd.read_excel(budget_file, sheet_name = 'QB', skiprows = 7)
    daily_spend = pd.read_excel(daily_spend_file, sheet_name = 'SEM Daily Spend SA360', skiprows=2)
    campaign_conditions = pd.read_excel(budget_file, sheet_name='Campaign_conditions')
    return (q_budget, daily_spend, campaign_conditions)

(q_budget, daily_spend, campaign_conditions) = load_inputs2()


# In[4]:


  ######    ##########  ##########          ######      ########    ########  ##########  ##########    ########  
  ######    ##########  ##########          ######      ########    ########  ##########  ##########    ########  
##          ##              ##            ##      ##  ##          ##          ##              ##      ##          
##          ##              ##            ##      ##  ##          ##          ##              ##      ##          
##    ####  ########        ##            ##########    ######      ######    ########        ##        ######    
##    ####  ########        ##            ##########    ######      ######    ########        ##        ######    
##      ##  ##              ##            ##      ##          ##          ##  ##              ##              ##  
##      ##  ##              ##            ##      ##          ##          ##  ##              ##              ##  
  ######    ##########      ##            ##      ##  ########    ########    ##########      ##      ########    
  ######    ##########      ##            ##      ##  ########    ########    ##########      ##      ########    



# In[5]:


sa360_lookup_cols = [
    'Campaign',
    'Headline 1',
    'Headline 2',
    'Description',
    'Description 1',
    'Description 2', 
    'Final URL'
]

sa360_lookup = (asset_file[sa360_lookup_cols]
    .drop_duplicates(subset=['Campaign'])
)


# In[6]:


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
# Keep only needed columns
daily_report = daily_spend[daily_cols]


# In[7]:


# ===============================
# DATA CLEANING
# ===============================

# Dates
daily_report['Day'] = pd.to_datetime(
    daily_report['Day'],
    dayfirst=True,
    errors='coerce'
)

# Numeric columns
num_cols = ['Clicks', 'Impr.']
daily_report[num_cols] = (
    daily_report[num_cols]
    .replace(',', '', regex=True)
    .apply(pd.to_numeric, errors='coerce')
)

# Money columns
money_cols = ['Avg. CPC', 'Cost + Fees']
daily_report[money_cols] = (
    daily_report[money_cols]
    .replace('[\$,]', '', regex=True)
    .astype(float)
)
daily_report['Cost'] = (
    daily_report['Cost + Fees'] / 1.0055
).round(2)

# ===============================
# MERGE (VLOOKUP STYLE)
# ===============================
sem_report = daily_report.merge(
    sa360_lookup,
    on='Campaign',
    how='left'
)

print('✅ SEM report ready with daily metrics + creatives')


# In[8]:


# ===============================
# CATEGORY ASSIGNMENT (EXCEL DRIVEN)
# ===============================

# 1. Asegurar el orden por prioridad si la columna existe en la tabla
if 'Priority' in campaign_conditions.columns:
    campaign_conditions = campaign_conditions.sort_values('Priority')

campaign = sem_report['Campaign'].fillna('')
sem_report['Partner'] = 'OPEX'  # Valor por defecto inicial

# 2. Iterar sobre cada regla definida en la hoja de Excel
for _, row in campaign_conditions.iterrows():
    # Nota: Aquí usamos la columna de categoría o partner de tu Excel (ajusta el nombre si en tu tabla se llama diferente, ej. 'Category')
    # Si la columna en tu Excel se llama 'Partner' pero aquí asigna 'Category', asegúrate de apuntar a la columna correcta de tu Excel.
    # Por seguridad, buscamos 'Category' o 'Partner' dependiendo de cómo nombraste la columna en tu archivo:
    target_col = 'Partner Campaign' if 'Partner Campaign' in campaign_conditions.columns else 'Partner'
    category_value = row[target_col]

    must_contain = str(row['Must_Contain']).strip() if pd.notna(row['Must_Contain']) else ''
    must_not_contain = str(row['Must_Not_Contain']).strip() if pd.notna(row['Must_Not_Contain']) else ''

    # Aplicar la regla únicamente a las filas que todavía están categorizadas como 'OPEX'
    mask = (sem_report['Partner'] == 'OPEX')

    # Evaluar palabras que DEBE contener (separadas por comas)
    if must_contain and must_contain.lower() != 'nan':
        keywords = [k.strip() for k in must_contain.split(',') if k.strip()]
        for kw in keywords:
            mask = mask & campaign.str.contains(kw, na=False, case=False)

    # Evaluar palabras que NO DEBE contener (separadas por comas)
    if must_not_contain and must_not_contain.lower() != 'nan':
        excl_keywords = [k.strip() for k in must_not_contain.split(',') if k.strip()]
        for kw in excl_keywords:
            mask = mask & (~campaign.str.contains(kw, na=False, case=False))

    # Asignar la categoría correspondiente a las filas que hicieron match
    sem_report.loc[mask, 'Partner'] = category_value

print("SEM Report Categories Assigned Successfully from Excel Rules")


# In[9]:


sem_report


# In[10]:


# =====================================================
# 8️⃣ DATE FILTER
# =====================================================

# Standardize daily spend dates
sem_report["Day"] = pd.to_datetime(
sem_report["Day"],
format="%Y-%m-%d",
errors="coerce")

# daily_spend["Day"] comes as YYYY-MM-DD
sem_report["Day"] = pd.to_datetime(
sem_report["Day"],
format="%Y-%m-%d",
errors="coerce")
# Convert dates
sem_report["Day"] = pd.to_datetime(
    daily_spend["Day"],
    dayfirst=True,
    errors="coerce")

q_budget["Partner ASD"] = pd.to_datetime(
    q_budget["Partner ASD"],
    errors="coerce")

q_budget["Partner AED"] = pd.to_datetime(
    q_budget["Partner AED"],
    errors="coerce")

# Create dictionaries
start_dates = q_budget.set_index("Partner")["Partner ASD"].to_dict()
end_dates = q_budget.set_index("Partner")["Partner AED"].to_dict()


# In[11]:


# Map dates by Partner
sem_report["Start Date"] = sem_report["Partner"].map(start_dates)
sem_report["End Date"] = sem_report["Partner"].map(end_dates)


# In[12]:


# Filter
sem_report = sem_report[(sem_report["Day"] >= sem_report["Start Date"]) &(sem_report["Day"] <= sem_report["End Date"])]


sem_report['Cost + Fees'] = pd.to_numeric(sem_report['Cost + Fees'], errors='coerce')


# In[ ]:





# In[13]:


# =====================================================
# 9️⃣ REPORT MONTH FILTER
# =====================================================

# Ask user for report month
report_month = input(
    "Enter report month (YYYY-MM), for example 2026-09: "
).strip()

# Validate input
try:
    report_period = pd.Period(report_month, freq="M")
except ValueError:
    raise ValueError(
        f"Invalid month '{report_month}'. "
        "Please use YYYY-MM format, for example: 2026-09"
    )

# Filter selected month
sem_report = sem_report[
    sem_report["Day"].dt.to_period("M") == report_period
].copy()

# Confirmation
print(f"\n📅 Generating report for: {report_period}")
print(f"📊 Rows included: {len(sem_report):,}")


# In[14]:


partner_periods = (
    sem_report
    .dropna(subset=["Day"])
    .groupby("Partner")["Day"]
    .agg(
        Start_Date="min",
        End_Date="max"
    )
    .to_dict(orient="index"))


# In[ ]:





# In[15]:


# ===============================
# AGGREGATE SEM REPORT (CUSTOM FIRST NON-EMPTY)
# ===============================

# 1. Definir columnas de agrupación
group_cols = [
    'Campaign', 'Partner', 'Category', 
    'Headline 1', 'Headline 2', 'Description', 'Description 1', 'Description 2', 'Final URL'
]
existing_group_cols = [col for col in group_cols if col in sem_report.columns]

# 2. Función auxiliar para encontrar el primer valor que no esté vacío ni sea '--'
def first_valid_text(series):
    # Filtrar valores que no sean nulos y que sean distintos de "--" (quitando espacios por si acaso)
    valid_values = series[
        pd.notna(series) & 
        (series.astype(str).str.strip() != '') & 
        (series.astype(str).str.strip() != '--')
    ]

    # Si encontramos alguno válido, regresamos el primero; si no, regresamos el valor original de la primera fila
    if not valid_values.empty:
        return valid_values.iloc[0]
    return series.iloc[0] if not series.empty else ''

# 3. Diccionario de agregación aplicando la función a los textos y 'sum' a las métricas
agg_dict = {
    'Clicks': 'sum',
    'Impr.': 'sum',
    'Cost + Fees': 'sum',
    'Cost': 'sum',
    # Aplicar nuestra función personalizada para evitar '--' o vacíos
    'Headline 1': first_valid_text,
    'Headline 2': first_valid_text,
    'Description': first_valid_text,
    'Description 1': first_valid_text,
    'Description 2': first_valid_text,
    'Final URL': first_valid_text
}

valid_agg_dict = {col: op for col, op in agg_dict.items() if col in sem_report.columns}

# 4. Ejecutar la agrupación
sem_report_aggregated = (
    sem_report
    .groupby(existing_group_cols, as_index=False)
    .agg(valid_agg_dict)
)

# 5. Recalcular métricas de tasa si es necesario
if 'Clicks' in sem_report_aggregated.columns and 'Impr.' in sem_report_aggregated.columns:
    sem_report_aggregated['CTR'] = np.where(
        sem_report_aggregated['Impr.'] > 0, 
        sem_report_aggregated['Clicks'] / sem_report_aggregated['Impr.'], 
        0
    )

if 'Cost + Fees' in sem_report_aggregated.columns and 'Clicks' in sem_report_aggregated.columns:
    sem_report_aggregated['Avg. CPC'] = np.where(
        sem_report_aggregated['Clicks'] > 0, 
        sem_report_aggregated['Cost + Fees'] / sem_report_aggregated['Clicks'], 
        0
    )

sem_report = sem_report_aggregated
print("SEM Report aggregated successfully, skipping '--' and blanks for text fields.")


# In[16]:


# Seleccionar solo las columnas necesarias para el resaltado
campaigns_to_highlight = sem_report[['Partner', 'Campaign']]

def export_campaigns_to_highlight(camp_df):
    # Get script location
    base_path = os.getcwd()

    # Create Output folder
    output_folder = os.path.join(base_path, 'Output')
    os.makedirs(output_folder, exist_ok=True)

    # Full file path para el CSV de campañas
    output_file = os.path.join(output_folder, 'Campaigns to Highlight.csv')

    # Guardar en CSV de forma limpia y eficiente
    camp_df.to_csv(output_file, index=False)

    print('Shazam! The Campaigns CSV is Ready!')

# RUN
export_campaigns_to_highlight(campaigns_to_highlight)


# In[17]:


# ===============================
# CHANNEL
# ===============================
c = sem_report['Campaign']
sem_report['Channel'] = np.select(
    [c.str.contains('Google'), c.str.contains('Bing')],
    ['Google', 'Bing'],
    default='Null'
)


# In[18]:


# ===============================
# FINAL COLUMNS DEFINITION
# ===============================
final_columns = [
    'Campaign',
    'Headline 1', 'Headline 2',
    'Description', 'Description 1', 'Description 2',
    'Clicks', 'Impr.', 'CTR', 'Avg. CPC',
    'Cost', 'Cost + Fees'
]

# ===============================
# SPLIT REPORTS & CLEAN COLUMNS
# ===============================
reports_separated = {
    f"{cat}_{ch}": df[final_columns].copy()
    for cat, df_cat in sem_report.groupby('Partner')
    for ch, df in df_cat.groupby('Channel')
}


# In[ ]:





# In[19]:


# =====================================================
# EXCEL REPORT GENERATION (DYNAMIC INPUTS)
# =====================================================

# Ask user for output folder name
folder_input = input(
    "Enter output folder name, for example 'September FY26 Monthly reports': "
).strip()

# Use a default folder name if the input is empty
if not folder_input:
    output_folder = "Default Name - Monthly reports"
else:
    output_folder = folder_input

os.makedirs(output_folder, exist_ok=True)

# =====================================================
# LOGO PATHS SETUP
# =====================================================
logos_folder = "Logos"
google_logo = os.path.join(logos_folder, "Google_Logo.png")
bing_logo = os.path.join(logos_folder, "Bing_Logo.png")

# Función auxiliar robusta para buscar el logo según el nombre del partner/motor
def get_logo_path(report_name):
    lower_name = report_name.lower()
    if 'bing' in lower_name:
        return bing_logo
    return google_logo # Por defecto usa el de Google o el genérico

for name, report in reports_separated.items():

    if report.empty:
        continue

    if "Cost + Fees" not in report.columns:
        continue

    total_cost = round(report['Cost + Fees'].sum(), 2)

    if total_cost == 0:
        continue

    total_cost_formatted = f"${total_cost:,.2f}"

    # Limpiar nombre del archivo para evitar caracteres extraños
    file_name = f"{name}_SEM_Total_Cost_{total_cost_formatted}_June_Q3_FY26.xlsx"
    file_path = os.path.join(output_folder, file_name)

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        money_format = workbook.add_format({'num_format': '$#,##0.00'})

        main_sheet_name = "Campaign Report"

        # ===============================
        # HEADER (MAIN TAB)
        # ===============================
        header_df = pd.DataFrame({
            'Report Name': [f'Partner: {name}'],
            'Period': ['June 1st to 30th Q3FY26']
        })
        header_df.to_excel(writer, sheet_name=main_sheet_name, index=False, startrow=0)

        # ===============================
        # MAIN REPORT
        # ===============================
        report.to_excel(writer, sheet_name=main_sheet_name, index=False, startrow=2)
        worksheet = writer.sheets[main_sheet_name]

        # ===============================
        # LOGO (MAIN TAB)
        # ===============================
        logo_path = get_logo_path(name)
        if os.path.exists(logo_path):
            worksheet.insert_image('E1', logo_path, {'x_scale': 0.5, 'y_scale': 0.4, 'x_offset': 100, 'y_offset': 5})
        else:
            print(f"⚠️ Logo not found: {logo_path}")

        # ===============================
        # COST BREAKDOWN TAB
        # ===============================
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

        # ===============================
        # FORMATTING (MONEY COLUMNS)
        # ===============================
        # Formato en Main Tab
        for col in ['Cost', 'Cost + Fees', 'Avg. CPC']:
            if col in report.columns:
                col_idx = report.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, None, money_format)

        # Formato en Cost Breakdown Tab (Columnas D y E usualmente Cost y Cost + Fees)
        cost_ws.set_column('D:E', None, money_format)

    # =================================================
    # INDIVIDUAL REPORT CONFIRMATION
    # =================================================
    print(f"✅ Report generated: {file_name} | Total Cost: {total_cost_formatted}")

# =====================================================
# FINAL CONFIRMATION
# =====================================================
print("\n✅ All reports are ready")
print(f"📁 Output folder: {output_folder}")


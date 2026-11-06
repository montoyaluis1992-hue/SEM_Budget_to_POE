#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import os

def get_file(filename):

    return os.path.join('Output',filename)

def load_inputs():

    campaigns_highlighter = get_file('Campaigns to Highlight.csv')

    # Validate
    for file in [campaigns_highlighter]:

        if not os.path.exists(file):

            raise FileNotFoundError(f'File not found: {file}')

    # Load Daily Spend csv
    campaigns_highlighter = pd.read_csv(campaigns_highlighter)

    print('Inputs Loaded')

    return (campaigns_highlighter)

# RUN
(campaigns_highlighter) = load_inputs()


# In[6]:


def get_invoice_csv(filename):
    return os.path.join('Invoice Input', filename)

def load_invoice_csv(): 
    # Definir las rutas de los archivos
    bing_path = get_invoice_csv('Invoice_Bing_SEM.xlsx')
    google_path = get_invoice_csv('Invoice_Google_SEM.csv')

    # Validar que ambos archivos existan antes de cargarlos
    for file in [bing_path, google_path]:
        if not os.path.exists(file):
            raise FileNotFoundError(f'File not found: {file}')

    # Preguntar interactivamente cuántas filas saltar para Google Invoice
    while True:
        try:
            skip_input = input("¿How many rows you need to skip in Google's Invoice? (10 by default): ").strip()

            # Si el usuario presiona Enter sin escribir nada, usamos 13 por defecto
            if not skip_input:
                skip_rows = 13
                break

            skip_rows = int(skip_input)
            break
        except ValueError:
            print("❌ Please use a valid number.")

    # Cargar los archivos usando las rutas y las filas dinámicas
    bing_invoice = pd.read_excel(bing_path, sheet_name='Invoice')
    google_invoice = pd.read_csv(google_path, skiprows=skip_rows)

    print(f'Invoices Loaded successfully (Google skipped {skip_rows} rows).')

    return bing_invoice, google_invoice

# Ejecutar la función
(bing_invoice, google_invoice) = load_invoice_csv()



# In[7]:


#########################
#####  Clean Data   #####
#########################

bing_invoice = bing_invoice.rename(columns={'Description':'Campaign','Total (USD)':'Spend'})
google_invoice = google_invoice[['Amount', 'Description']]
google_invoice = google_invoice.rename(columns={'Amount':'Spend', 'Description':'Campaign'})
google_invoice['Spend'] = pd.to_numeric(google_invoice['Spend'].astype(str).str.replace(',', ''), errors='coerce')
bing_invoice['Campaign'] = bing_invoice['Campaign'].str.replace(r'[\r\n]+', '', regex=True).str.strip()


# In[8]:


import re

# Aplicar limpieza solo a las celdas que comienzan con "Invalid Activity"
google_invoice['Campaign'] = google_invoice['Campaign'].apply(
    lambda x: re.sub(r'^Invalid\s*Activity.*?(US_.+)', r'\1', x, flags=re.IGNORECASE)
    if isinstance(x, str) and x.strip().lower().startswith('invalid activity')
    else x
)
invoices_csv = pd.concat([google_invoice, bing_invoice], ignore_index=True)
invoices_csv = invoices_csv[['Spend', 'Campaign']]
#print(invoices_csv.head())
campaigns_highlighter = campaigns_highlighter.rename(columns={'Partner Category': 'Partner'})


# In[9]:


#print(campaigns_highlighter.head())

sem_partner_list = sorted(campaigns_highlighter['Partner'].unique())


# In[11]:


#Create an empty DF for each category
for x in sem_partner_list: 
    globals()[x] = pd.DataFrame()

#Extract the lines into each DF
for x in sem_partner_list: 
    df_categ = campaigns_highlighter[campaigns_highlighter['Partner'] == x].copy()
    globals()[x] = pd.concat([globals()[x], df_categ])


# In[12]:


for x in sem_partner_list: 
    df = globals()[x]
    print((f"Number of rows in {x}: {df.shape[0]}"))


# In[13]:


# Iterar sobre los DataFrames de la lista sem_reports_list
for df_name in sem_partner_list:
    # Obtener el DataFrame por su nombre de la lista
    df = globals()[df_name]

    # Agrupar por 'Campaign' y 'Category' y agregar la columna 'Account' como listas
    grouped_df = df.groupby(['Campaign', 'Partner']).apply(list).reset_index()

    # Actualizar el DataFrame original con el DataFrame agrupado
    globals()[df_name] = grouped_df

    # Obtener el número de líneas del DataFrame agrupado
    num_lines = len(grouped_df)

    # Mostrar el número de líneas del DataFrame actual
    #print(f"Número de líneas para {df_name}: {num_lines}")


# In[15]:


# Create empty DataFrames for each category
for category in sem_partner_list:
    globals()[category + '_Google'] = pd.DataFrame()
    globals()[category + '_Bing'] = pd.DataFrame()

# List to store the names of new DataFrames
new_dataframe_names = []

# Iterate over the list and split each DataFrame
for category in sem_partner_list:
    df = globals()[category]
    google_df = df[df['Campaign'].str.contains('Google', case=False, na=False)].copy()
    bing_df = df[df['Campaign'].str.contains('Bing', case=False, na=False)].copy()
    globals()[category + '_Google'] = google_df
    globals()[category + '_Bing'] = bing_df

    # Add names of new DataFrames to the list (excluding those containing 'OPEX')
    if not google_df.empty:
        if 'OPEX' not in category:
            new_dataframe_names.append(category + '_Google')
        else:
            del globals()[category + '_Google']  # Delete DataFrame containing 'OPEX'

    if not bing_df.empty:
        if 'OPEX' not in category:
            new_dataframe_names.append(category + '_Bing')
        else:
            del globals()[category + '_Bing']  # Delete DataFrame containing 'OPEX'

    # Delete DataFrames if they are empty
    if google_df.empty:
        del globals()[category + '_Google']
    if bing_df.empty:
        del globals()[category + '_Bing']

# Print the number of rows for each new DataFrame
for category in sem_partner_list:
    if category + '_Google' in globals():
        print(f"Number of rows in {category}_Google: {globals()[category + '_Google'].shape[0]}")
    if category + '_Bing' in globals():
        print(f"Number of rows in {category}_Bing: {globals()[category + '_Bing'].shape[0]}")

# Print the list of names of new DataFrames
#print("List of names of new DataFrames:")
#print(new_dataframe_names)


# In[17]:


import os
import fitz  # PyMuPDF

def highlight_campaigns(pdf_path, df_campaigns, output_pdf_name, output_folder):
    if not os.path.exists(pdf_path):
        print(f"⚠️ The PDF file does not exist: {pdf_path}")
        return

    pdf_document = fitz.open(pdf_path)
    if pdf_document.page_count == 0:
        print(f"⚠️ the doc {pdf_path} is empty.")
        pdf_document.close()
        return

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]

        for campaign in df_campaigns:
            campaign = str(campaign).strip()
            if not campaign:
                continue

            # --- BÚSQUEDA ESTRICTA Y EXACTA ---
            # Buscamos todas las apariciones de la campaña en la página
            found_rects = page.search_for(campaign)

            matches = []
            for rect in found_rects:
                # Extraemos el texto exacto que encierra este rectángulo en el PDF
                snippet = page.get_textbox(rect).strip()

                # VALIDACIÓN CRUCIAL: 
                # Verificamos que el texto del PDF sea exactamente igual a la campaña, 
                # evitando así que tome una versión más larga (ej. que tenga "_Bestseller" al final).
                if snippet.lower() == campaign.lower():
                    matches.append(rect)

            # --- Aplicar el resaltado transparente y ajustado a los matches validados ---
            for rect in matches:
                # 📐 Ajuste de dimensiones: más alto y ligeramente más delgado
                adjusted_rect = fitz.Rect(
                    rect.x0 + 1.5,  # Margen izquierdo
                    rect.y0 - 1.0,  # Margen superior (más alto)
                    rect.x1 - 1.5,  # Margen derecho
                    rect.y1 + 1.0   # Margen inferior (más alto)
                )

                # Crear la anotación de resaltado translúcida
                highlight = page.add_highlight_annot(adjusted_rect)
                highlight.update()

    # Crear la carpeta de salida si no existe
    os.makedirs(output_folder, exist_ok=True)

    # Guardar PDF resultante en la carpeta especificada
    output_path = os.path.join(output_folder, f'{output_pdf_name}.pdf')
    pdf_document.save(output_path)
    pdf_document.close()
    print(f"✅ Invoice successfully generated!!!: {output_pdf_name}.pdf")

# =====================================================
# CONFIGURACIÓN DE CARPETA DE SALIDA (DINÁMICO)
# =====================================================
folder_input = input(
    "Enter output folder name suffix, for example 'September Q4FY26': "
).strip()

if not folder_input:
    folder_suffix = "September FY26"
else:
    folder_suffix = folder_input

# El nombre de la carpeta llevará el prefijo solicitado
output_folder = f"Highlighted Invoices SEM - {folder_suffix}"


# =====================================================
# LOOP PRINCIPAL
# =====================================================
for df_name in new_dataframe_names:
    if df_name in globals():
        df = globals()[df_name]

        # Unir con los costos (invoices_csv)
        df_with_cost = df.merge(invoices_csv, on='Campaign', how='left')
        df_with_cost["Spend"] = df_with_cost["Spend"].fillna(0)

        # Sumar el costo total
        total_cost = df_with_cost["Spend"].sum()
        cost_str = f"${total_cost:,.0f}"

        # Crear nombre del PDF usando únicamente el sufijo limpio (ej. September Q4FY26) y el costo
        output_pdf_name = f'{df_name}_SEM_{folder_suffix}_Invoice_{cost_str}'

        # Definir rutas de los PDFs dentro de la carpeta "Invoice Input"
        google_pdf_path = os.path.join("Invoice Input", "Invoice_Google_SEM.pdf")
        bing_pdf_path = os.path.join("Invoice Input", "Invoice_Bing_SEM.pdf")

        # Ejecutar highlight según fuente
        if 'Google' in df_name:
            highlight_campaigns(google_pdf_path, df['Campaign'], output_pdf_name, output_folder)
        elif 'Bing' in df_name:
            highlight_campaigns(bing_pdf_path, df['Campaign'], output_pdf_name, output_folder)

print(f"\n✅ Process Completed. The new PDFs are ready in the '{output_folder}' folder.")


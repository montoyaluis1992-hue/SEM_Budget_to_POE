#!/usr/bin/env python
# coding: utf-8

import os
import re
import traceback
import warnings
import fitz  # PyMuPDF
import numpy as np
import pandas as pd

# Ignorar advertencias de obsolescencia de librerías para mantener la consola limpia
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=DeprecationWarning)


# =====================================================
# LOAD ASSETS & INVOICES
# =====================================================

def get_file(filename):
    return os.path.join('Output', filename)

def load_inputs():
    campaigns_highlighter = get_file('Campaigns to Highlight.csv')
    if not os.path.exists(campaigns_highlighter):
        raise FileNotFoundError(f'File not found: {campaigns_highlighter}')

    df = pd.read_csv(campaigns_highlighter)
    print('Inputs Loaded')
    return df

def load_invoice_csv():
    bing_path = os.path.join('Invoice Input', 'Invoice_Bing_SEM.xlsx')
    google_path = os.path.join('Invoice Input', 'Invoice_Google_SEM.csv')

    for file in [bing_path, google_path]:
        if not os.path.exists(file):
            raise FileNotFoundError(f'File not found: {file}')

    while True:
        try:
            skip_input = input("¿How many rows you need to skip in Google's Invoice? (10 by default): ").strip()
            if not skip_input:
                skip_rows = 13
                break
            skip_rows = int(skip_input)
            break
        except ValueError:
            print("❌ Please use a valid number.")

    bing_invoice = pd.read_excel(bing_path, sheet_name='Invoice')
    google_invoice = pd.read_csv(google_path, skiprows=skip_rows)

    print(f'Invoices Loaded successfully (Google skipped {skip_rows} rows).')
    return bing_invoice, google_invoice


# =====================================================
# PDF HIGHLIGHT FUNCTION
# =====================================================

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

            found_rects = page.search_for(campaign)

            matches = []
            for rect in found_rects:
                snippet = page.get_textbox(rect).strip()
                if snippet.lower() == campaign.lower():
                    matches.append(rect)

            for rect in matches:
                adjusted_rect = fitz.Rect(
                    rect.x0 + 1.5,
                    rect.y0 - 1.0,
                    rect.x1 - 1.5,
                    rect.y1 + 1.0
                )
                highlight = page.add_highlight_annot(adjusted_rect)
                highlight.update()

    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, f'{output_pdf_name}.pdf')
    pdf_document.save(output_path)
    pdf_document.close()
    print(f"✅ Invoice successfully generated!!!: {output_pdf_name}.pdf")


# =====================================================
# MAIN PROCESS
# =====================================================

def main():
    campaigns_highlighter = load_inputs()
    bing_invoice, google_invoice = load_invoice_csv()

    # ===============================
    # CLEAN DATA (Using .loc to avoid warnings)
    # ===============================
    bing_invoice = bing_invoice.rename(columns={'Description': 'Campaign', 'Total (USD)': 'Spend'}).copy()
    google_invoice = google_invoice[['Amount', 'Description']].copy()
    google_invoice = google_invoice.rename(columns={'Amount': 'Spend', 'Description': 'Campaign'})
    
    google_invoice.loc[:, 'Spend'] = pd.to_numeric(
        google_invoice['Spend'].astype(str).str.replace(',', '', regex=False),
        errors='coerce'
    )
    bing_invoice.loc[:, 'Campaign'] = bing_invoice['Campaign'].str.replace(r'[\r\n]+', '', regex=True).str.strip()

    # Aplicar limpieza solo a las celdas que comienzan con "Invalid Activity"
    google_invoice.loc[:, 'Campaign'] = google_invoice['Campaign'].apply(
        lambda x: re.sub(r'^Invalid\s*Activity.*?(US_.+)', r'\1', x, flags=re.IGNORECASE)
        if isinstance(x, str) and x.strip().lower().startswith('invalid activity')
        else x
    )

    invoices_csv = pd.concat([google_invoice, bing_invoice], ignore_index=True)
    invoices_csv = invoices_csv[['Spend', 'Campaign']].copy()
    
    campaigns_highlighter = campaigns_highlighter.rename(columns={'Partner Category': 'Partner'}).copy()

    sem_partner_list = sorted(campaigns_highlighter['Partner'].dropna().unique())

    # Create dynamic dataframes dictionary to replace globals safely
    partner_dfs = {}

    for x in sem_partner_list:
        df_categ = campaigns_highlighter[campaigns_highlighter['Partner'] == x].copy()
        partner_dfs[x] = df_categ

    for x in sem_partner_list:
        df = partner_dfs[x]
        print(f"Number of rows in {x}: {df.shape[0]}")

    for df_name in sem_partner_list:
        df = partner_dfs[df_name]
        grouped_df = df.groupby(['Campaign', 'Partner']).apply(list).reset_index()
        partner_dfs[df_name] = grouped_df

    # Split into Google and Bing
    new_dataframe_names = []

    for category in sem_partner_list:
        df = partner_dfs[category]
        google_df = df[df['Campaign'].str.contains('Google', case=False, na=False)].copy()
        bing_df = df[df['Campaign'].str.contains('Bing', case=False, na=False)].copy()

        if not google_df.empty and 'OPEX' not in category:
            partner_dfs[category + '_Google'] = google_df
            new_dataframe_names.append(category + '_Google')

        if not bing_df.empty and 'OPEX' not in category:
            partner_dfs[category + '_Bing'] = bing_df
            new_dataframe_names.append(category + '_Bing')

    for category in sem_partner_list:
        g_name = category + '_Google'
        b_name = category + '_Bing'
        if g_name in partner_dfs:
            print(f"Number of rows in {g_name}: {partner_dfs[g_name].shape[0]}")
        if b_name in partner_dfs:
            print(f"Number of rows in {b_name}: {partner_dfs[b_name].shape[0]}")

    # =====================================================
    # CONFIGURACIÓN DE CARPETA DE SALIDA
    # =====================================================
    folder_input = input(
        "Enter output folder name suffix, for example 'September Q4FY26': "
    ).strip()

    if not folder_input:
        folder_suffix = "September FY26"
    else:
        folder_suffix = folder_input

    output_folder = f"Highlighted Invoices SEM - {folder_suffix}"

    # =====================================================
    # LOOP PRINCIPAL
    # =====================================================
    for df_name in new_dataframe_names:
        if df_name in partner_dfs:
            df = partner_dfs[df_name]

            df_with_cost = df.merge(invoices_csv, on='Campaign', how='left').copy()
            df_with_cost.loc[:, "Spend"] = df_with_cost["Spend"].fillna(0)

            total_cost = df_with_cost["Spend"].sum()
            cost_str = f"${total_cost:,.0f}"

            output_pdf_name = f'{df_name}_SEM_{folder_suffix}_Invoice_{cost_str}'

            google_pdf_path = os.path.join("Invoice Input", "Invoice_Google_SEM.pdf")
            bing_pdf_path = os.path.join("Invoice Input", "Invoice_Bing_SEM.pdf")

            if 'Google' in df_name:
                highlight_campaigns(google_pdf_path, df['Campaign'], output_pdf_name, output_folder)
            elif 'Bing' in df_name:
                highlight_campaigns(bing_pdf_path, df['Campaign'], output_pdf_name, output_folder)

    print(f"\n✅ Process Completed. The new PDFs are ready in the '{output_folder}' folder.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n[ERROR] Something went wrong:")
        traceback.print_exc()
    finally:
        input("\nClick ENTER to close this window...")

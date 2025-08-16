import pandas as pd
import unicodedata
import os
import datetime
import re
input_folder = r'C:\Users\user\Downloads\code nouveau'
hl_df = pd.read_excel(os.path.join(input_folder,'HL_MATERIAUX.xlsx'))
siren_naf_df = pd.read_excel(os.path.join(input_folder,'SIREN_APE.xlsx'),
dtype={'Code APE': str})
naf_df = pd.read_excel(os.path.join(input_folder,'CF_WF_NAF_France_2024_Adjusted-code.xlsx'))

#Step 2 — Clean Supplier Names & Merge with Identifiers
def clean_name(name):
    if isinstance(name, str):
        name = re.sub(r'$E$','', name)
        name = unicodedata.normalize('NFKD', name).encode('ASCII','ignore').decode('utf-8')
        name = " ".join(name.split()).lower()
        return name.strip()
    return name

hl_df['Fournisseurs_Eiffage'] = hl_df['Fournisseur enfant panel'].apply(clean_name)
siren_df['Fournisseurs_Eiffage'] = siren_df['Fournisseur'].apply(clean_name)

merged_df = pd.merge(
    hl_df,
siren_df[['Fournisseurs_Eiffage','Code SIREN','Code NAF']],
on='Fournisseurs_Eiffage',
how='left'
)

#Step 3 — Manual Completion of Missing Codes
missing_rows = merged_df[merged_df['Code SIREN'].isnull() |
merged_df['Code NAF'].isnull()]
for idx, row in missing_rows.iterrows():
supplier_name = row['Fournisseur enfant panel']
print(f"\nSupplier not found: {supplier_name}")
    while True:
code_siren = input(f"SIREN (9 digits): ").strip()
    if len(code_siren) == 9 and code_siren.isdigit():
break
    while True:
code_ape = input(f"NAF (4 digits + 1 letter): ").strip().upper()
    if re.match(r'^\d{4}[A-Z]$'
    , code_naf):
break
merged_df.loc[idx,'Code SIREN'] = code_siren
merged_df.loc[idx,'Code APE'] = code_naf
#Step 4 — Harmonize APE and NAF Codes for Sector Mapping
merged_df['Code NAF Clean'] = merged_df['CodeNAF'].str.replace(r'\.','', regex=True).str.strip()
naf_df['Code NAF Clean'] = naf_df['CodeNAF'].str.replace(r'\.','', regex=True).str.strip()

merged_sector = pd.merge(
merged_df[['Panel parent','Panel enfant','Fournisseurenfant panel','Code SIREN','Code NAF','Code NAFClean']],naf_df[['Code NAF Clean','new best match sector','kgCO2-eq per EUR2024']],
left_on='Code APE Clean'
,
right_on='Code NAF Clean'
,
how='left'
)
#Step 6 — Calculate Environmental Impacts
def safe_multiply(x, y):
try:
return float(x) * float(y)
except (TypeError, ValueError):
return None
merged_sector['GHG Emissions (kg CO2)'] = merged_sector.apply(
lambda row: safe_multiply(row['DEPENSES'], row['kg CO2-eq per
EUR2024']), axis=1)
Step 7 — Structure Final Report for Export
structured_rows = []
grouped = merged_sector.groupby(['Panel parent'
,
'Panel enfant'])
for (panel_parent, panel_enfant), group in grouped:
total_depenses = group['DEPENSES'].sum(skipna=True)
structured_rows.append({
'Panel parent': panel_parent,
'Panel enfant': panel_enfant,
'Fournisseur': ""
,
'DÉPENSES (€)': f"Total : {total_depenses:,.0f}".replace(",","").replace(".",","),'Code NAF': "",'Code SIREN': "",'Émissions CO2 (kg)': ""})
for _, row in group.iterrows():
structured_rows.append({
'Panel parent': panel_parent,
'Panel enfant': panel_enfant,
'Fournisseur': row['Fournisseur enfant panel'][:35],
'DÉPENSES (€)': f"{row['DEPENSES']:,.0f}".replace("
,
"
,
" ").replace("."
,
"
,
"),
'Code NAF': row['Code NAF'],
'Code SIREN': row['Code SIREN'],
'Émissions CO2 (kg)': f"{row['GHG Emissions (kg CO2)']:,.0f}".replace("
,
"
,
" ").replace("."
,
"
,
") if pd.notna(row['GHG Emissions (kg CO2)']) else ""
,
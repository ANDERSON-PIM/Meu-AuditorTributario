import streamlit as st
import pandas as pd
import csv
import os
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image

CONFIG_FILE = 'configuracoes.csv'

# Configuração da página
st.set_page_config(page_title="Sistema de Auditoria Tributária - Escritório Contábil Sigilo", layout="centered")

# ✅ Exibe a logo corretamente
try:
    logo = Image.open("logo.png")
    st.image(logo, width=250)
except FileNotFoundError:
    st.warning("⚠️ A logo 'logo.png' não foi encontrada. Coloque na mesma pasta do app.py.")

# Título
st.title("📊 Sistema de Auditoria Tributária de Produtos - Escritório Contábil Sigilo")

# === Funções ===

def get_keywords(text):
    stopwords = {'de', 'da', 'do', 'e', 'em', 'com', 'ml'}
    return set(word.lower() for word in str(text).split() if word.lower() not in stopwords and len(word) > 2)

def load_configurations():
    configs = {}
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) == 4:
                    configs[row[0]] = {
                        'ncm': row[1],
                        'aliq_icms': row[2],
                        'tributacao': row[3]
                    }
    return configs

def save_all_configurations(configs):
    with open(CONFIG_FILE, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        for desc, values in configs.items():
            writer.writerow([desc, values['ncm'], values['aliq_icms'], values['tributacao']])

def process_planilha(df, configs):
    if 'TRIBUTAÇÃO' not in df.columns:
        df['TRIBUTAÇÃO'] = None

    df['NCM Alterado'] = False
    df['Aliq. ICMS Alterado'] = False
    df['TRIBUTAÇÃO Alterado'] = False

    for i, row in df.iterrows():
        desc_item = str(row['Descrição item']).strip().lower()
        ncm_item = str(row['NCM']).strip()
        aliq_item = str(row['Aliq. ICMS']).strip()
        trib_item = str(row['TRIBUTAÇÃO']).strip() if row['TRIBUTAÇÃO'] else ''

        palavras_item = get_keywords(desc_item)
        encontrado = False

        for desc_base, values in configs.items():
            palavras_base = get_keywords(desc_base)
            palavras_iguais = palavras_item & palavras_base

            if len(palavras_iguais) >= 2 or values['ncm'] == ncm_item:
                encontrado = True

                if ncm_item != values['ncm']:
                    df.at[i, 'NCM'] = values['ncm']
                    df.at[i, 'NCM Alterado'] = True
                if aliq_item != values['aliq_icms']:
                    df.at[i, 'Aliq. ICMS'] = values['aliq_icms']
                    df.at[i, 'Aliq. ICMS Alterado'] = True
                if trib_item != values['tributacao']:
                    df.at[i, 'TRIBUTAÇÃO'] = values['tributacao']
                    df.at[i, 'TRIBUTAÇÃO Alterado'] = True
                break
    return df

def aplicar_destaque_excel(df, output_file):
    df.to_excel(output_file, index=False)

    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    col_map = {col: idx + 1 for idx, col in enumerate(df.columns)}

    for i in range(2, len(df) + 2):
        if df.at[i - 2, 'NCM Alterado']:
            ws.cell(row=i, column=col_map['NCM']).fill = yellow_fill
        if df.at[i - 2, 'Aliq. ICMS Alterado']:
            ws.cell(row=i, column=col_map['Aliq. ICMS']).fill = yellow_fill
        if df.at[i - 2, 'TRIBUTAÇÃO Alterado']:
            ws.cell(row=i, column=col_map['TRIBUTAÇÃO']).fill = yellow_fill

    wb.save(output_file)

# === Execução do App ===

configs = load_configurations()

# 1. Adicionar base auditada
st.header("🟩 1. Adicionar base auditada")
uploaded_base = st.file_uploader("Envie a planilha auditada", type=['xlsx'], key='base')
if uploaded_base:
    base_df = pd.read_excel(uploaded_base)
    base_df.columns = base_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

    for _, row in base_df.iterrows():
        desc = str(row['Descrição item']).strip()
        ncm = str(row['NCM']).strip()
        aliq = str(row['Aliq. ICMS']).strip()
        trib = str(row.get('TRIBUTAÇÃO', aliq)).strip()
        configs[desc] = {'ncm': ncm, 'aliq_icms': aliq, 'tributacao': trib}

    save_all_configurations(configs)
    st.success("✅ Base auditada adicionada com sucesso!")

# 2. Ver base auditada com filtro
st.header("📋 2. Ver base de configurações salva (com filtro)")
search_term = st.text_input("🔎 Pesquise por NCM ou parte da descrição")
if search_term:
    search_term = search_term.lower().strip()
    df_base = pd.DataFrame.from_dict(configs, orient='index')
    df_base = df_base.reset_index().rename(columns={'index': 'Descrição'})
    df_filtrada = df_base[
        df_base['Descrição'].str.lower().str.contains(search_term) |
        df_base['ncm'].str.lower().str.contains(search_term)
    ]
    st.dataframe(df_filtrada)
elif st.checkbox("👁️ Mostrar toda a base auditada"):
    df_base = pd.DataFrame.from_dict(configs, orient='index')
    df_base = df_base.reset_index().rename(columns={'index': 'Descrição'})
    st.dataframe(df_base)

# 3. Auditoria
st.header("🔍 3. Auditar nova planilha")
uploaded_audit = st.file_uploader("Envie a planilha para auditoria", type=['xlsx'], key='audit')
if uploaded_audit:
    audit_df = pd.read_excel(uploaded_audit)
    audit_df.columns = audit_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

    result_df = process_planilha(audit_df, configs)
    st.success("✅ Planilha auditada com sucesso!")

    output_file = "resultado_auditoria.xlsx"
    aplicar_destaque_excel(result_df, output_file)

    with open(output_file, 'rb') as f:
        st.download_button("📥 Baixar resultado auditado", f, file_name=output_file)

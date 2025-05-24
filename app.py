import streamlit as st
import pandas as pd
import csv
import os
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from thefuzz import fuzz  # Import para fuzzy matching

CONFIG_FILE = 'configuracoes.csv'

st.set_page_config(page_title="Sistema de Auditoria Tributária - Escritório Contábil Sigilo", layout="centered")

try:
    logo = Image.open("logo.png")
    st.image(logo, width=250)
except FileNotFoundError:
    st.warning("⚠️ A logo 'logo.png' não foi encontrada. Coloque na mesma pasta do app.py.")

st.title("📊 Sistema de Auditoria Tributária de Produtos - Escritório Contábil Sigilo")

def get_keywords(text):
    stopwords = {'de', 'da', 'do', 'e', 'em', 'com', 'ml'}
    return set(word.lower() for word in str(text).split() if word.lower() not in stopwords and len(word) > 2)

def load_configurations():
    configs = {}
    # Verificar se existe o arquivo CSV
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                sample = f.read(1024)  # Ler uma amostra para determinar o formato
                dialect = csv.Sniffer().sniff(sample)
                f.seek(0)  # Voltar ao início do arquivo
                
                reader = csv.reader(f, dialect)
                for row in reader:
                    if len(row) >= 3:  # Garantir que temos pelo menos descrição, NCM e alíquota
                        ncm_val = str(row[1]).strip() # Garantir que NCM seja string
                        aliq_val = str(row[2]).strip()
                        if len(row) == 4:
                            # Formato: [descrição, NCM, alíquota, tributação]
                            configs[row[0]] = {
                                'NCM': ncm_val,
                                'ALIQ_ICMS': aliq_val,
                                'TRIBUTACAO': str(row[3]).strip(),
                                'CEST': '0'  # Valor padrão para CEST
                            }
                        elif len(row) == 5:
                            # Formato: [descrição, NCM, alíquota, tributação, CEST]
                            configs[row[0]] = {
                                'NCM': ncm_val,
                                'ALIQ_ICMS': aliq_val,
                                'TRIBUTACAO': str(row[3]).strip(),
                                'CEST': str(row[4]).strip()
                            }
            
            st.sidebar.success(f"✅ Arquivo CSV carregado com sucesso! {len(configs)} itens encontrados.")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar arquivo CSV: {str(e)}")
    
    # Se não existir arquivo CSV ou estiver vazio, tentar carregar do Excel
    if len(configs) == 0 and os.path.exists('configuracoes.xlsx'):
        try:
            # Ler Excel garantindo que NCM seja string
            df = pd.read_excel('configuracoes.xlsx', dtype={'NCM': str})
            df.columns = df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)
            
            for _, row in df.iterrows():
                desc = str(row['Descrição item']).strip()
                # NCM já é string devido ao dtype, apenas tratar valores ausentes
                ncm = str(row['NCM']).strip() if pd.notna(row['NCM']) else '' 
                aliq = str(row['Aliq. ICMS']).strip()
                trib = str(row.get('TRIBUTAÇÃO', aliq)).strip()
                cest = str(row.get('CEST', '0')).strip()  # Atribuindo 0 se não tiver
                configs[desc] = {'NCM': ncm, 'ALIQ_ICMS': aliq, 'TRIBUTACAO': trib, 'CEST': cest}
            
            # Salvar em CSV para uso futuro
            save_all_configurations(configs)
            st.sidebar.success(f"✅ Base carregada do arquivo Excel 'configuracoes.xlsx'. {len(configs)} itens encontrados.")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar arquivo Excel: {str(e)}")
    
    return configs

def save_all_configurations(configs):
    with open(CONFIG_FILE, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        for desc, values in configs.items():
            # Garantir que NCM seja salvo como string
            writer.writerow([desc, str(values['NCM']), str(values['ALIQ_ICMS']), str(values['TRIBUTACAO']), str(values['CEST'])])

def aplicar_destaque_excel(df, output_file):
    # Garantir que NCM seja string antes de salvar
    if 'NCM' in df.columns:
        df['NCM'] = df['NCM'].astype(str)
        
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
        if df.at[i - 2, 'CEST Alterado']:
            ws.cell(row=i, column=col_map['CEST']).fill = yellow_fill

    wb.save(output_file)

def export_to_pdf(df, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica", 10)
    c.drawString(30, height - 30, "Relatório de Auditoria Tributária")

    c.drawString(30, height - 50, "Descrição")
    c.drawString(200, height - 50, "NCM")
    c.drawString(350, height - 50, "Aliq. ICMS")
    c.drawString(500, height - 50, "TRIBUTAÇÃO")
    c.drawString(650, height - 50, "CEST")  # Coluna CEST

    y_position = height - 70
    for i, row in df.iterrows():
        c.drawString(30, y_position, str(row['Descrição item']))
        # Garantir que NCM seja exibido como string
        c.drawString(200, y_position, str(row['NCM'])) 
        c.drawString(350, y_position, str(row['Aliq. ICMS']))
        c.drawString(500, y_position, str(row['TRIBUTAÇÃO']))
        c.drawString(650, y_position, str(row['CEST']))  # Exibindo CEST
        y_position -= 15

        if y_position < 40:
            c.showPage()
            c.setFont("Helvetica", 10)
            y_position = height - 30
    c.save()

def process_planilha(df, configs):
    # Se não existe a coluna 'CEST', adicionar com valor '0' por padrão
    if 'CEST' not in df.columns:
        df['CEST'] = '0'

    # Garantir que NCM seja tratado como string logo após a leitura
    if 'NCM' in df.columns:
        df['NCM'] = df['NCM'].astype(str).str.replace('\.0$', '', regex=True).str.strip()
    else:
        df['NCM'] = '' # Criar coluna NCM vazia se não existir
        
    # Garantir que outras colunas relevantes sejam string
    if 'Aliq. ICMS' in df.columns: df['Aliq. ICMS'] = df['Aliq. ICMS'].astype(str).str.strip()
    if 'TRIBUTAÇÃO' in df.columns: df['TRIBUTAÇÃO'] = df['TRIBUTAÇÃO'].astype(str).str.strip()
    if 'CEST' in df.columns: df['CEST'] = df['CEST'].astype(str).str.strip()

    df['NCM Alterado'] = False
    df['Aliq. ICMS Alterado'] = False
    df['TRIBUTAÇÃO Alterado'] = False
    df['CEST Alterado'] = False
    df['ITEM CONSIDERADO'] = ''

    for i, row in df.iterrows():
        desc_item = str(row['Descrição item']).strip().lower()
        # NCM já deve ser string aqui
        ncm_item = str(row['NCM']).strip() 
        aliq_item = str(row['Aliq. ICMS']).strip()
        trib_item = str(row['TRIBUTAÇÃO']).strip() if pd.notna(row['TRIBUTAÇÃO']) else ''
        cest_item = str(row['CEST']).strip() if pd.notna(row['CEST']) else ''

        palavras_item = get_keywords(desc_item)
        encontrado = False

        # 1) Correspondência exata da descrição
        for desc_base, values in configs.items():
            desc_base_clean = desc_base.strip().lower()
            # Garantir que NCM da base também seja string para comparação
            ncm_base = str(values['NCM']).strip()
            aliq_base = str(values['ALIQ_ICMS']).strip()
            trib_base = str(values['TRIBUTACAO']).strip()
            cest_base = str(values['CEST']).strip()
            
            if desc_item == desc_base_clean:
                if ncm_item != ncm_base:
                    df.at[i, 'NCM'] = ncm_base
                    df.at[i, 'NCM Alterado'] = True
                if aliq_item != aliq_base:
                    df.at[i, 'Aliq. ICMS'] = aliq_base
                    df.at[i, 'Aliq. ICMS Alterado'] = True
                if trib_item != trib_base:
                    df.at[i, 'TRIBUTAÇÃO'] = trib_base
                    df.at[i, 'TRIBUTAÇÃO Alterado'] = True
                if cest_item != cest_base:
                    df.at[i, 'CEST'] = cest_base
                    df.at[i, 'CEST Alterado'] = True
                df.at[i, 'ITEM CONSIDERADO'] = f'Código: {desc_base}'
                encontrado = True
                break
        if encontrado:
            continue

        # 2) Pelo menos 2 palavras em comum ou NCM igual
        for desc_base, values in configs.items():
            palavras_base = get_keywords(desc_base)
            palavras_iguais = palavras_item & palavras_base
            ncm_base = str(values['NCM']).strip()
            aliq_base = str(values['ALIQ_ICMS']).strip()
            trib_base = str(values['TRIBUTACAO']).strip()
            cest_base = str(values['CEST']).strip()

            if len(palavras_iguais) >= 2 or ncm_base == ncm_item:
                if ncm_item != ncm_base:
                    df.at[i, 'NCM'] = ncm_base
                    df.at[i, 'NCM Alterado'] = True
                if aliq_item != aliq_base:
                    df.at[i, 'Aliq. ICMS'] = aliq_base
                    df.at[i, 'Aliq. ICMS Alterado'] = True
                if trib_item != trib_base:
                    df.at[i, 'TRIBUTAÇÃO'] = trib_base
                    df.at[i, 'TRIBUTAÇÃO Alterado'] = True
                if cest_item != cest_base:
                    df.at[i, 'CEST'] = cest_base
                    df.at[i, 'CEST Alterado'] = True
                df.at[i, 'ITEM CONSIDERADO'] = f'Código: {desc_base}'
                encontrado = True
                break
        if encontrado:
            continue

        # 3) Fuzzy matching (limite 70%)
        for desc_base, values in configs.items():
            score = fuzz.ratio(desc_item, desc_base.strip().lower())
            ncm_base = str(values['NCM']).strip()
            aliq_base = str(values['ALIQ_ICMS']).strip()
            trib_base = str(values['TRIBUTACAO']).strip()
            cest_base = str(values['CEST']).strip()
            
            if score >= 70:
                if ncm_item != ncm_base:
                    df.at[i, 'NCM'] = ncm_base
                    df.at[i, 'NCM Alterado'] = True
                if aliq_item != aliq_base:
                    df.at[i, 'Aliq. ICMS'] = aliq_base
                    df.at[i, 'Aliq. ICMS Alterado'] = True
                if trib_item != trib_base:
                    df.at[i, 'TRIBUTAÇÃO'] = trib_base
                    df.at[i, 'TRIBUTAÇÃO Alterado'] = True
                if cest_item != cest_base:
                    df.at[i, 'CEST'] = cest_base
                    df.at[i, 'CEST Alterado'] = True
                df.at[i, 'ITEM CONSIDERADO'] = f'Código: {desc_base}'
                encontrado = True
                break

    return df

# Carregar configurações
configs = load_configurations()

# Informações de status na barra lateral
st.sidebar.write("### Status do Sistema")
st.sidebar.write(f"Itens na base de configurações: {len(configs)}")
if len(configs) == 0:
    st.sidebar.warning("⚠️ Base de configurações vazia. Por favor, adicione uma base auditada ou verifique se o arquivo 'configuracoes.csv' ou 'configuracoes.xlsx' está na mesma pasta do app.")
else:
    st.sidebar.success("✅ Base de configurações carregada com sucesso!")
    # Mostrar exemplo de item para debug
    try:
        exemplo_key = list(configs.keys())[0]
        st.sidebar.write(f"Exemplo de item: {exemplo_key}")
        st.sidebar.write(f"Valores: {configs[exemplo_key]}")
    except IndexError:
        st.sidebar.write("Não foi possível mostrar exemplo (base vazia).")

tab1, tab2, tab3 = st.tabs(["🟩 1. Adicionar Base Auditada", "📋 2. Ver Base de Configurações", "🔍 3. Auditoria"])

with tab1:
    st.header("1. Adicionar base auditada")
    uploaded_base = st.file_uploader("Envie a planilha auditada", type=['xlsx'], key='base')
    if uploaded_base:
        # Ler Excel garantindo que NCM seja string
        base_df = pd.read_excel(uploaded_base, dtype={'NCM': str})
        base_df.columns = base_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

        # Garantir que 'CEST' esteja presente na planilha auditada
        if 'CEST' not in base_df.columns:
            base_df['CEST'] = '0'

        for _, row in base_df.iterrows():
            desc = str(row['Descrição item']).strip()
            # NCM já é string devido ao dtype, apenas tratar valores ausentes
            ncm = str(row['NCM']).strip() if pd.notna(row['NCM']) else ''
            aliq = str(row['Aliq. ICMS']).strip()
            trib = str(row.get('TRIBUTAÇÃO', aliq)).strip()
            cest = str(row.get('CEST', '0')).strip()
            configs[desc] = {'NCM': ncm, 'ALIQ_ICMS': aliq, 'TRIBUTACAO': trib, 'CEST': cest}

        save_all_configurations(configs)
        st.success(f"✅ Base auditada adicionada com sucesso! {len(configs)} itens na base.")
        # Atualizar status na sidebar
        st.sidebar.success("✅ Base de configurações atualizada!")
        st.sidebar.write(f"Itens na base de configurações: {len(configs)}")

with tab2:
    st.header("2. Ver base de configurações salva (com filtro)")
    
    if len(configs) == 0:
        st.warning("⚠️ A base de configurações está vazia. Por favor, adicione uma base auditada primeiro ou verifique se o arquivo 'configuracoes.csv' ou 'configuracoes.xlsx' está na mesma pasta do app.")
    else:
        search_term = st.text_input("🔎 Pesquise por NCM ou parte da descrição")
        
        if search_term or st.checkbox("👁️ Mostrar toda a base auditada"):
            # Criar DataFrame a partir do dicionário configs
            df_base = pd.DataFrame.from_dict(configs, orient='index')
            
            # Garantir que NCM seja string no DataFrame ANTES de exibir
            if 'NCM' in df_base.columns:
                df_base['NCM'] = df_base['NCM'].astype(str)
            
            # Verificar e mostrar as colunas disponíveis
            st.write(f"Colunas disponíveis: {df_base.columns.tolist()}")
            
            # Renomear colunas
            df_base = df_base.reset_index().rename(columns={'index': 'Descrição'})
            df_base['Descrição'] = df_base['Descrição'].astype(str)
            
            # Garantir que NCM seja string após renomear (redundante, mas seguro)
            if 'NCM' in df_base.columns:
                df_base['NCM'] = df_base['NCM'].astype(str)
                
            df_display = df_base.copy() # Criar cópia para exibição
            
            if search_term:
                search_term = search_term.lower().strip()
                # Filtrar por descrição
                filtro_descricao = df_display['Descrição'].str.lower().str.contains(search_term)
                
                # Filtrar por NCM se a coluna existir
                if 'NCM' in df_display.columns:
                    # NCM já é string, apenas comparar
                    filtro_ncm = df_display['NCM'].str.lower().str.contains(search_term)
                    df_filtrada = df_display[filtro_descricao | filtro_ncm]
                else:
                    st.warning("Coluna NCM não encontrada. Filtrando apenas por descrição.")
                    df_filtrada = df_display[filtro_descricao]
                
                # Exibir o DataFrame filtrado
                st.dataframe(df_filtrada)
            else:
                # Exibir o DataFrame completo
                st.dataframe(df_display)

with tab3:
    st.header("3. Auditoria")
    
    if len(configs) == 0:
        st.warning("⚠️ A base de configurações está vazia. Por favor, adicione uma base auditada primeiro ou verifique se o arquivo 'configuracoes.csv' ou 'configuracoes.xlsx' está na mesma pasta do app.")
    else:
        uploaded_audit = st.file_uploader("Envie a planilha para auditoria", type=['xlsx'], key='audit')
        if uploaded_audit:
            # Ler Excel garantindo que NCM seja string
            audit_df = pd.read_excel(uploaded_audit, dtype={'NCM': str})
            audit_df.columns = audit_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

            result_df = process_planilha(audit_df.copy(), configs) # Usar cópia para evitar modificar original
            st.success("✅ Planilha auditada com sucesso!")

            output_excel_file = "resultado_auditoria.xlsx"
            aplicar_destaque_excel(result_df, output_excel_file)

            with open(output_excel_file, 'rb') as f:
                st.download_button("📥 Baixar resultado auditado (Excel)", f, file_name=output_excel_file)

            output_pdf_file = "resultado_auditoria.pdf"
            export_to_pdf(result_df, output_pdf_file)

            with open(output_pdf_file, 'rb') as f:
                st.download_button("📥 Baixar resultado auditado (PDF)", f, file_name=output_pdf_file)

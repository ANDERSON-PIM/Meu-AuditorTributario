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

# Carregar logo (se presente)
try:
    logo = Image.open("logo.png")
    st.image(logo, width=250)
except FileNotFoundError:
    st.warning("⚠️ A logo 'logo.png' não foi encontrada. Coloque na mesma pasta do app.py.")

st.title("📊 Sistema de Auditoria Tributária de Produtos - Escritório Contábil Sigilo")

def clean_cest(cest_value):
    """Limpa o valor do CEST, removendo '.0' e garantindo que seja uma string."""
    if pd.isna(cest_value):
        return '0'
    cest_str = str(cest_value).strip()
    if cest_str.endswith('.0'):
        cest_str = cest_str[:-2] # Remove '.0'
    # Tenta converter para int para remover quaisquer outros decimais e depois para str
    try:
        return str(int(cest_str))
    except ValueError:
        # Se não for um número válido após remover .0, retorna a string limpa
        # Isso pode acontecer se o CEST original for algo como 'INVALIDO.0'
        return cest_str if cest_str else '0'

def get_keywords(text):
    stopwords = {'de', 'da', 'do', 'e', 'em', 'com', 'ml'}
    return set(word.lower() for word in str(text).split() if word.lower() not in stopwords and len(word) > 2)

def load_configurations():
    configs = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                # Tenta detectar o dialeto (delimitador, etc.)
                try:
                    sample = f.read(2048) # Ler uma amostra maior pode ajudar
                    dialect = csv.Sniffer().sniff(sample)
                    f.seek(0)
                except csv.Error:
                    # Se Sniffer falhar, tenta com delimitador padrão (vírgula)
                    f.seek(0)
                    dialect = csv.excel # Usa um dialeto padrão
                    st.sidebar.warning("Não foi possível detectar o delimitador do CSV automaticamente, tentando com vírgula.")

                reader = csv.reader(f, dialect)
                header = next(reader, None) # Pula o cabeçalho se existir

                for row in reader:
                    if len(row) >= 5: # Espera 5 colunas: Desc, NCM, ALIQ, TRIB, CEST
                        desc_val = str(row[0]).strip()
                        ncm_val = str(row[1]).strip()
                        aliq_val = str(row[2]).strip()
                        trib_val = str(row[3]).strip()
                        cest_val = clean_cest(row[4]) # Limpa o CEST lido do CSV
                        configs[desc_val] = {
                            'NCM': ncm_val,
                            'ALIQ_ICMS': aliq_val,
                            'TRIBUTACAO': trib_val,
                            'CEST': cest_val
                        }
                    elif len(row) == 4: # Caso antigo sem CEST explícito
                         desc_val = str(row[0]).strip()
                         ncm_val = str(row[1]).strip()
                         aliq_val = str(row[2]).strip()
                         trib_val = str(row[3]).strip()
                         configs[desc_val] = {
                            'NCM': ncm_val,
                            'ALIQ_ICMS': aliq_val,
                            'TRIBUTACAO': trib_val,
                            'CEST': '0' # Assume CEST 0 se não presente
                        }

            st.sidebar.success(f"✅ Arquivo CSV '{CONFIG_FILE}' carregado com sucesso! {len(configs)} itens encontrados.")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar arquivo CSV '{CONFIG_FILE}': {str(e)}")

    # Se o CSV estava vazio ou não existe, tenta carregar do Excel
    if not configs and os.path.exists('configuracoes.xlsx'):
        st.sidebar.info("Arquivo CSV não encontrado ou vazio. Tentando carregar de 'configuracoes.xlsx'...")
        try:
            # Especifica dtype para NCM e CEST para evitar conversão automática para float
            df = pd.read_excel('configuracoes.xlsx', dtype={'NCM': str, 'CEST': str})
            df.columns = df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

            # Verifica se as colunas essenciais existem
            required_cols = ['Descrição item', 'NCM', 'Aliq. ICMS']
            if not all(col in df.columns for col in required_cols):
                st.sidebar.error(f"Erro: O arquivo Excel deve conter as colunas: {', '.join(required_cols)}")
                return {}

            for _, row in df.iterrows():
                desc = str(row['Descrição item']).strip()
                ncm = str(row['NCM']).strip() if pd.notna(row['NCM']) else ''
                aliq = str(row['Aliq. ICMS']).strip()
                # Usa .get para TRIBUTAÇÃO e CEST para caso não existam
                trib = str(row.get('TRIBUTAÇÃO', aliq)).strip() # Usa Aliq. ICMS como fallback para TRIBUTAÇÃO
                cest = clean_cest(row.get('CEST', '0')) # Limpa o CEST lido do Excel

                if desc: # Só adiciona se a descrição não for vazia
                    configs[desc] = {'NCM': ncm, 'ALIQ_ICMS': aliq, 'TRIBUTACAO': trib, 'CEST': cest}

            if configs:
                save_all_configurations(configs) # Salva no formato CSV padrão
                st.sidebar.success(f"✅ Base carregada do arquivo Excel 'configuracoes.xlsx' e salva em '{CONFIG_FILE}'. {len(configs)} itens encontrados.")
            else:
                 st.sidebar.warning("Nenhum item válido encontrado no arquivo Excel 'configuracoes.xlsx'.")

        except Exception as e:
            st.sidebar.error(f"Erro ao carregar arquivo Excel 'configuracoes.xlsx': {str(e)}")

    return configs

def save_all_configurations(configs):
    """Salva todas as configurações no arquivo CSV, garantindo que CEST esteja limpo."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            # Escreve o cabeçalho
            writer.writerow(['Descrição item', 'NCM', 'Aliq. ICMS', 'TRIBUTAÇÃO', 'CEST'])
            for desc, values in configs.items():
                # Garante que todos os valores sejam strings e CEST esteja limpo
                ncm = str(values.get('NCM', '')).strip()
                aliq = str(values.get('ALIQ_ICMS', '')).strip()
                trib = str(values.get('TRIBUTACAO', '')).strip()
                cest = clean_cest(values.get('CEST', '0')) # Garante limpeza ao salvar
                writer.writerow([desc, ncm, aliq, trib, cest])
        # st.sidebar.info(f"Configurações salvas em {CONFIG_FILE}") # Opcional: feedback de salvamento
    except Exception as e:
        st.sidebar.error(f"Erro ao salvar configurações em '{CONFIG_FILE}': {str(e)}")

def process_planilha(df, configs):
    """Processa a planilha de auditoria comparando com as configurações."""
    # Garante a existência e limpeza inicial das colunas no DataFrame de entrada
    if 'NCM' not in df.columns:
        df['NCM'] = ''
    else:
        df['NCM'] = df['NCM'].astype(str).str.replace('\.0$', '', regex=True).str.strip()

    if 'Aliq. ICMS' not in df.columns:
        df['Aliq. ICMS'] = ''
    else:
        df['Aliq. ICMS'] = df['Aliq. ICMS'].astype(str).str.strip()

    if 'TRIBUTAÇÃO' not in df.columns:
        df['TRIBUTAÇÃO'] = ''
    else:
        df['TRIBUTAÇÃO'] = df['TRIBUTAÇÃO'].astype(str).str.strip()

    if 'CEST' not in df.columns:
        df['CEST'] = '0'
    # Limpa a coluna CEST do DataFrame de entrada usando a função clean_cest
    df['CEST'] = df['CEST'].apply(clean_cest)

    # Inicializa colunas de controle
    df['NCM Alterado'] = False
    df['Aliq. ICMS Alterado'] = False
    df['TRIBUTAÇÃO Alterado'] = False
    df['CEST Alterado'] = False
    df['ITEM CONSIDERADO'] = ''
    df['SIMILARIDADE'] = 0.0 # Adiciona coluna para score de similaridade

    # Itera sobre cada linha da planilha de auditoria
    for i, row in df.iterrows():
        desc_item = str(row.get('Descrição item', '')).strip().lower()
        if not desc_item:
            continue # Pula linhas sem descrição

        # Obtém valores da linha atual (já limpos na etapa inicial)
        ncm_item = str(row.get('NCM', '')).strip()
        aliq_item = str(row.get('Aliq. ICMS', '')).strip()
        trib_item = str(row.get('TRIBUTAÇÃO', '')).strip()
        cest_item = str(row.get('CEST', '0')).strip() # Já deve estar limpo

        palavras_item = get_keywords(desc_item)
        melhor_match = None
        max_score = -1
        match_type = "Nenhum"

        # 1. Procura por correspondência exata na descrição
        if desc_item in configs:
             melhor_match = desc_item
             max_score = 100 # Score máximo para correspondência exata
             match_type = "Descrição Exata"

        # 2. Se não encontrou exata, procura por palavras-chave ou NCM
        if not melhor_match:
            for desc_base, values in configs.items():
                desc_base_clean = desc_base.strip().lower()
                palavras_base = get_keywords(desc_base_clean)
                palavras_iguais = palavras_item & palavras_base
                ncm_base = str(values.get('NCM', '')).strip()

                # Considera match se tiver >= 2 palavras iguais OU NCM igual (e não vazio)
                if len(palavras_iguais) >= 2 or (ncm_base and ncm_base == ncm_item):
                    # Calcula um score baseado nas palavras e NCM para desempate
                    score_atual = len(palavras_iguais) * 10 + (50 if ncm_base and ncm_base == ncm_item else 0)
                    if score_atual > max_score:
                        max_score = score_atual
                        melhor_match = desc_base
                        match_type = "Palavras/NCM"

        # 3. Se ainda não encontrou, usa fuzzy matching
        if not melhor_match:
            for desc_base, values in configs.items():
                desc_base_clean = desc_base.strip().lower()
                score = fuzz.ratio(desc_item, desc_base_clean)
                if score >= 70 and score > max_score: # Usa 70 como limite e pega o maior score
                    max_score = score
                    melhor_match = desc_base
                    match_type = f"Similaridade ({score}%)"

        # Se encontrou um melhor match por qualquer método
        if melhor_match:
            valores_base = configs[melhor_match]
            ncm_base = str(valores_base.get('NCM', '')).strip()
            aliq_base = str(valores_base.get('ALIQ_ICMS', '')).strip()
            trib_base = str(valores_base.get('TRIBUTACAO', '')).strip()
            cest_base = clean_cest(valores_base.get('CEST', '0')) # Garante limpeza do CEST da base

            # Compara e atualiza os campos, marcando as alterações
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

            df.at[i, 'ITEM CONSIDERADO'] = f'{match_type}: {melhor_match}'
            df.at[i, 'SIMILARIDADE'] = max_score # Salva o score do match
        else:
            df.at[i, 'ITEM CONSIDERADO'] = 'Nenhuma correspondência encontrada'
            df.at[i, 'SIMILARIDADE'] = 0

    return df

def aplicar_destaque_excel(df, filename):
    """Salva o DataFrame em Excel com destaque nas células alteradas."""
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado Auditoria')
            workbook = writer.book
            worksheet = writer.sheets['Resultado Auditoria']

            # Define o estilo de preenchimento (amarelo claro)
            yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')

            # Encontra os índices das colunas pelo nome
            cols = {col_name: idx + 1 for idx, col_name in enumerate(df.columns)}

            # Itera pelas linhas do DataFrame (começa da linha 2 no Excel por causa do cabeçalho)
            for row_idx, row_data in df.iterrows():
                excel_row = row_idx + 2 # +1 para 1-based index, +1 para pular cabeçalho

                if row_data.get('NCM Alterado', False):
                    if 'NCM' in cols:
                        worksheet.cell(row=excel_row, column=cols['NCM']).fill = yellow_fill
                if row_data.get('Aliq. ICMS Alterado', False):
                    if 'Aliq. ICMS' in cols:
                        worksheet.cell(row=excel_row, column=cols['Aliq. ICMS']).fill = yellow_fill
                if row_data.get('TRIBUTAÇÃO Alterado', False):
                    if 'TRIBUTAÇÃO' in cols:
                        worksheet.cell(row=excel_row, column=cols['TRIBUTAÇÃO']).fill = yellow_fill
                if row_data.get('CEST Alterado', False):
                    if 'CEST' in cols:
                        worksheet.cell(row=excel_row, column=cols['CEST']).fill = yellow_fill

            # Remove as colunas de controle 'Alterado' e 'SIMILARIDADE' antes de salvar (opcional)
            # Se quiser manter as colunas de controle no Excel, comente as linhas abaixo
            # df_final = df.drop(columns=['NCM Alterado', 'Aliq. ICMS Alterado', 'TRIBUTAÇÃO Alterado', 'CEST Alterado', 'SIMILARIDADE'], errors='ignore')
            # df_final.to_excel(writer, index=False, sheet_name='Resultado Auditoria')
            # Se remover as colunas, o destaque precisa ser aplicado antes da remoção ou ajustar os índices

        st.info(f"Arquivo Excel '{filename}' gerado com destaque.")
    except Exception as e:
        st.error(f"Erro ao gerar arquivo Excel com destaque: {str(e)}")

def export_to_pdf(df, filename):
    """Exporta o DataFrame para um arquivo PDF simples."""
    try:
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter
        margin = 50
        y_position = height - margin
        line_height = 12

        # Cabeçalho do PDF
        c.setFont("Helvetica-Bold", 10)
        x_offset = margin
        col_widths = {} # Ajustar conforme necessário ou calcular dinamicamente

        # Desenha cabeçalho da tabela
        header_names = [col for col in df.columns if not col.endswith('Alterado') and col != 'SIMILARIDADE']
        # Define larguras (exemplo simples, pode precisar de ajuste) 
        num_cols = len(header_names)
        default_width = (width - 2 * margin) / num_cols
        for i, col_name in enumerate(header_names):
             col_widths[col_name] = default_width
             c.drawString(x_offset + i * default_width, y_position, col_name)
        y_position -= line_height * 1.5

        # Desenha linhas da tabela
        c.setFont("Helvetica", 8)
        for _, row in df.iterrows():
            if y_position < margin + line_height:
                c.showPage()
                c.setFont("Helvetica-Bold", 10)
                y_position = height - margin
                # Redesenha cabeçalho na nova página
                for i, col_name in enumerate(header_names):
                    c.drawString(x_offset + i * default_width, y_position, col_name)
                y_position -= line_height * 1.5
                c.setFont("Helvetica", 8)

            x_offset_current = margin
            for i, col_name in enumerate(header_names):
                cell_value = str(row.get(col_name, ''))
                # Simplificação: Truncar texto longo
                max_len = int(col_widths[col_name] / 4) # Estimativa de caracteres
                display_text = (cell_value[:max_len] + '...') if len(cell_value) > max_len else cell_value
                c.drawString(x_offset_current + i * default_width, y_position, display_text)

            y_position -= line_height

        c.save()
        st.info(f"Arquivo PDF '{filename}' gerado.")
    except Exception as e:
        st.error(f"Erro ao gerar arquivo PDF: {str(e)}")

# --- Interface Streamlit --- 

# Carregar configurações iniciais
configs = load_configurations()

# Informações de status na barra lateral
st.sidebar.write("### Status do Sistema")
st.sidebar.write(f"Itens na base de configurações: {len(configs)}")
if not configs:
    st.sidebar.warning("⚠️ Base de configurações vazia. Adicione uma base na Aba 1 ou verifique os arquivos 'configuracoes.csv'/'configuracoes.xlsx'.")
else:
    st.sidebar.success("✅ Base de configurações carregada.")

# Criação das abas no Streamlit
tab1, tab2, tab3 = st.tabs(["🟩 1. Adicionar/Atualizar Base", "📋 2. Ver Base de Configurações", "🔍 3. Realizar Auditoria"])

# Aba 1 - Adicionar/Atualizar Base Auditada
with tab1:
    st.header("1. Adicionar ou Atualizar Base de Configurações")
    st.markdown("Envie uma planilha Excel (`.xlsx`) auditada para adicionar novos itens ou atualizar existentes na base de configurações.")
    uploaded_base = st.file_uploader("Selecione a planilha auditada", type=['xlsx'], key='base_uploader')

    if uploaded_base:
        try:
            # Lê a planilha enviada, tratando NCM e CEST como string
            base_df = pd.read_excel(uploaded_base, dtype={'NCM': str, 'CEST': str})
            base_df.columns = base_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

            # Verifica colunas essenciais
            required_cols = ['Descrição item', 'NCM', 'Aliq. ICMS']
            if not all(col in base_df.columns for col in required_cols):
                 st.error(f"Erro: A planilha enviada deve conter as colunas: {', '.join(required_cols)}")
            else:
                itens_adicionados = 0
                itens_atualizados = 0
                # Itera sobre a planilha enviada
                for _, row in base_df.iterrows():
                    desc = str(row['Descrição item']).strip()
                    if not desc: continue # Pula linhas sem descrição

                    ncm = str(row['NCM']).strip() if pd.notna(row['NCM']) else ''
                    aliq = str(row['Aliq. ICMS']).strip()
                    trib = str(row.get('TRIBUTAÇÃO', aliq)).strip() # Usa Aliq. ICMS como fallback
                    cest = clean_cest(row.get('CEST', '0')) # Limpa o CEST

                    # Verifica se o item já existe para contar como atualização
                    if desc in configs:
                        itens_atualizados += 1
                    else:
                        itens_adicionados += 1

                    # Adiciona ou atualiza na memória
                    configs[desc] = {'NCM': ncm, 'ALIQ_ICMS': aliq, 'TRIBUTACAO': trib, 'CEST': cest}

                # Salva todas as configurações (incluindo as novas/atualizadas) no CSV
                save_all_configurations(configs)
                st.success(f"✅ Base atualizada com sucesso! Itens adicionados: {itens_adicionados}, Itens atualizados: {itens_atualizados}. Total na base: {len(configs)}.")
                # Atualiza o status na sidebar
                st.sidebar.write(f"Itens na base de configurações: {len(configs)}")
                st.rerun() # Recarrega a página para refletir a base atualizada nas outras abas

        except Exception as e:
            st.error(f"Erro ao processar a planilha enviada: {str(e)}")

# Aba 2 - Visualizar Base de Configurações
with tab2:
    st.header("2. Ver Base de Configurações Salva")
    st.markdown("Visualize e pesquise os itens atualmente na base de configurações.")

    if not configs:
        st.warning("⚠️ A base de configurações está vazia.")
    else:
        search_term = st.text_input("🔎 Pesquisar por Descrição, NCM ou CEST", key='search_base')
        show_all = st.checkbox("👁️ Mostrar toda a base", key='show_all_base')

        # Cria DataFrame a partir do dicionário configs para exibição
        if configs:
            df_base = pd.DataFrame.from_dict(configs, orient='index')
            df_base = df_base.reset_index().rename(columns={'index': 'Descrição item'})
            # Garante a ordem das colunas
            display_cols = ['Descrição item', 'NCM', 'CEST', 'Aliq. ICMS', 'TRIBUTAÇÃO']
            df_display = df_base[[col for col in display_cols if col in df_base.columns]]
        else:
            df_display = pd.DataFrame(columns=['Descrição item', 'NCM', 'CEST', 'Aliq. ICMS', 'TRIBUTAÇÃO'])

        # Aplica filtro se houver termo de pesquisa
        if search_term:
            search_term_lower = search_term.lower().strip()
            mask = df_display.apply(lambda row: any(search_term_lower in str(cell).lower() for cell in row), axis=1)
            df_filtrada = df_display[mask]
            st.dataframe(df_filtrada, use_container_width=True)
            st.caption(f"{len(df_filtrada)} itens encontrados para '{search_term}'.")
        elif show_all:
            st.dataframe(df_display, use_container_width=True)
            st.caption(f"Mostrando todos os {len(df_display)} itens da base.")
        else:
            st.info("Digite um termo de pesquisa ou marque 'Mostrar toda a base' para ver os dados.")

# Aba 3 - Realizar Auditoria
with tab3:
    st.header("3. Realizar Auditoria")
    st.markdown("Envie uma planilha Excel (`.xlsx`) para ser auditada com base nas configurações atuais.")

    if not configs:
        st.warning("⚠️ A base de configurações está vazia. Adicione uma base na Aba 1 primeiro.")
    else:
        uploaded_audit = st.file_uploader("Selecione a planilha para auditoria", type=['xlsx'], key='audit_uploader')

        if uploaded_audit:
            try:
                # Lê a planilha de auditoria, tratando NCM e CEST como string
                audit_df = pd.read_excel(uploaded_audit, dtype={'NCM': str, 'CEST': str})
                audit_df.columns = audit_df.columns.str.strip().str.replace('"', '', regex=False).str.replace('\n', '', regex=False).str.replace('\r', '', regex=False)

                # Verifica se a coluna 'Descrição item' existe
                if 'Descrição item' not in audit_df.columns:
                    st.error("Erro: A planilha de auditoria deve conter a coluna 'Descrição item'.")
                else:
                    st.info("Processando auditoria... Isso pode levar alguns instantes.")
                    # Processa a planilha
                    result_df = process_planilha(audit_df.copy(), configs)
                    st.success("✅ Auditoria concluída com sucesso!")

                    # Prepara arquivos para download
                    output_excel_file = "resultado_auditoria.xlsx"
                    output_pdf_file = "resultado_auditoria.pdf"

                    # Gera Excel com destaque
                    aplicar_destaque_excel(result_df.copy(), output_excel_file)
                    # Gera PDF (versão simplificada)
                    export_to_pdf(result_df.copy(), output_pdf_file)

                    # Mostra prévia do resultado (opcional, primeiras 50 linhas)
                    st.dataframe(result_df.head(50), use_container_width=True)
                    st.caption("Prévia das primeiras 50 linhas do resultado.")

                    # Botões de download
                    col1, col2 = st.columns(2)
                    with col1:
                        try:
                            with open(output_excel_file, 'rb') as f_excel:
                                st.download_button(
                                    label="📥 Baixar Resultado (Excel)",
                                    data=f_excel,
                                    file_name=output_excel_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        except FileNotFoundError:
                            st.error(f"Erro: Não foi possível encontrar o arquivo {output_excel_file} para download.")

                    with col2:
                         try:
                            with open(output_pdf_file, 'rb') as f_pdf:
                                st.download_button(
                                    label="📥 Baixar Resultado (PDF)",
                                    data=f_pdf,
                                    file_name=output_pdf_file,
                                    mime="application/pdf"
                                )
                         except FileNotFoundError:
                             st.error(f"Erro: Não foi possível encontrar o arquivo {output_pdf_file} para download.")

            except Exception as e:
                st.error(f"Erro ao processar a auditoria: {str(e)}")


import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import unicodedata
import zipfile
import io
import os
import re

# Configuração da página Web
st.set_page_config(page_title="Organizador de CT-e", page_icon="📦", layout="wide")

st.title("📦 Organizador Automático de CT-e")
st.markdown("Faça o upload dos arquivos abaixo para organizar os XMLs em pastas e gerar o relatório.")

# ==========================================================
# FUNÇÕES DE TRATAMENTO
# ==========================================================

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize("NFD", str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()

def extrair_apenas_digitos_cte(valor):
    if pd.isna(valor):
        return ""
    val_str = str(valor).split('-')[0].split('.')[0].strip()
    digitos = re.sub(r'\D', '', val_str)
    if digitos:
        return str(int(digitos))
    return ""

def determinar_produto_e_tes(descricao, uf_origem, uf_destino, df_regras):
    desc_norm = normalizar_texto(descricao)
    eh_intraestadual = (str(uf_origem).strip().upper() == str(uf_destino).strip().upper())

    if not df_regras.empty:
        df_regras_ordenadas = df_regras.copy()
        df_regras_ordenadas["TAM_DESC"] = df_regras_ordenadas["DESC_NORM"].str.len()
        df_regras_ordenadas = df_regras_ordenadas.sort_values(by="TAM_DESC", ascending=False)

        for _, row in df_regras_ordenadas.iterrows():
            desc_regra = row["DESC_NORM"]
            if desc_regra in desc_norm:
                return row["Conta_Str"], (row["TES_Intra"] if eh_intraestadual else row["TES_Inter"])

            palavras_chave = [w for w in desc_regra.split() if w not in ["DE", "PARA", "P/", "EM", "DO", "POR", "CONTA"]]
            if len(palavras_chave) > 0 and all(kw in desc_norm for kw in palavras_chave):
                return row["Conta_Str"], (row["TES_Intra"] if eh_intraestadual else row["TES_Inter"])

    if "DEVOLUCAO" in desc_norm:
        return "51063", ("052" if eh_intraestadual else "054")
    elif "GARANTIA" in desc_norm or "TROCA" in desc_norm:
        return "51061", ("054" if eh_intraestadual else "052")
    elif "TRANSFERENCIA" in desc_norm or "TRANSF" in desc_norm:
        return "51064", ("054" if eh_intraestadual else "052")
    elif "VENDA" in desc_norm:
        return "28197", ("045" if eh_intraestadual else "044")
    
    return "NAO_MAPEADO", "000"

# ==========================================================
# INTERFACE DO STREAMLIT (UPLOADS)
# ==========================================================

col1, col2 = st.columns(2)

with col1:
    arquivo_excel_ctes = st.file_uploader("1. Planilha de CT-es (excel.xlsx)", type=["xlsx"])
    arquivo_contas = st.file_uploader("2. Planilha de Regras (Contas.xlsx)", type=["xlsx"])

with col2:
    zip_xmls = st.file_uploader("3. Pasta compactada com XMLs (.zip)", type=["zip"])

# Se a planilha for enviada, permite selecionar a coluna do CT-e manualmente
coluna_cte_selecionada = None
if arquivo_excel_ctes:
    try:
        df_temp = pd.read_excel(arquivo_excel_ctes, nrows=5)
        df_temp.columns = df_temp.columns.str.strip()
        
        # Sugere coluna padrão
        idx_padrao = 0
        for i, col in enumerate(df_temp.columns):
            if any(term in col.upper() for term in ["CTE", "NCTE", "DOC", "CHAVE"]):
                idx_padrao = i
                break

        coluna_cte_selecionada = st.selectbox(
            "🎯 Selecione a coluna da planilha que contém o NÚMERO DO CT-e ou CHAVE DE ACESSO:",
            options=list(df_temp.columns),
            index=idx_padrao
        )
    except Exception:
        pass

# ==========================================================
# BOTÃO DE PROCESSAMENTO
# ==========================================================

if st.button("🚀 Processar e Organizar CT-es", type="primary", use_container_width=True):
    if not arquivo_excel_ctes or not arquivo_contas or not zip_xmls:
        st.error("⚠️ Por favor, faça o upload dos 3 arquivos para continuar!")
    else:
        with st.spinner("Processando arquivos... Aguarde um instante."):
            try:
                # 1. Carrega Regras da Planilha Contas
                df_regras = pd.read_excel(arquivo_contas, sheet_name="Conta")
                df_regras["DESC_NORM"] = df_regras["Descrição"].apply(normalizar_texto)
                df_regras["Conta_Str"] = df_regras["Conta"].astype(str).str.split('.').str[0].str.strip()
                df_regras["TES_Intra"] = df_regras["Intraestadual"].astype(str).str.split('.').str[0].str.strip().str.zfill(3)
                df_regras["TES_Inter"] = df_regras["interestadual"].astype(str).str.split('.').str[0].str.strip().str.zfill(3)

                # 2. Ler todos os XMLs
                xml_por_numero = {}
                xml_por_chave = {}
                amostra_xmls = []

                with zipfile.ZipFile(zip_xmls, 'r') as z:
                    for filename in z.namelist():
                        if filename.lower().endswith('.xml') and not filename.startswith('__MACOSX'):
                            try:
                                content = z.read(filename)
                                root = ET.fromstring(content)

                                numero, cfop, uf_origem, uf_destino, chave = None, None, None, None, None

                                for elem in root.iter():
                                    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                                    if tag == "nCT":
                                        numero = elem.text.strip() if elem.text else None
                                    elif tag == "CFOP":
                                        cfop = elem.text.strip() if elem.text else None
                                    elif tag == "UFIni":
                                        uf_origem = elem.text.strip() if elem.text else None
                                    elif tag == "UFFim":
                                        uf_destino = elem.text.strip() if elem.text else None
                                    elif tag in ["chCTe", "infCte"]:
                                        # Pega a chave de acesso de 44 dígitos
                                        if "Id" in elem.attrib:
                                            chave = re.sub(r'\D', '', elem.attrib["Id"])
                                        elif elem.text and len(re.sub(r'\D', '', elem.text)) == 44:
                                            chave = re.sub(r'\D', '', elem.text)

                                dados_xml = {
                                    "filename": os.path.basename(filename),
                                    "content": content,
                                    "cfop": cfop or "",
                                    "origem": uf_origem or "",
                                    "destino": uf_destino or ""
                                }

                                if numero:
                                    num_limpo = extrair_apenas_digitos_cte(numero)
                                    if num_limpo:
                                        xml_por_numero[num_limpo] = dados_xml
                                        if len(amostra_xmls) < 5:
                                            amostra_xmls.append(f"Número nCT: {num_limpo}")

                                if chave:
                                    xml_por_chave[chave] = dados_xml

                            except Exception:
                                pass

                # 3. Ler Planilha
                df_ctes = pd.read_excel(arquivo_excel_ctes)
                df_ctes.columns = df_ctes.columns.str.strip()

                coluna_cte = coluna_cte_selecionada or df_ctes.columns[0]

                amostra_planilha = [
                    f"Valor na coluna '{coluna_cte}': {val}" 
                    for val in df_ctes[coluna_cte].dropna().head(5)
                ]

                coluna_desc = next((c for c in df_ctes.columns if "DESC" in c.upper()), df_ctes.columns[1] if len(df_ctes.columns) > 1 else df_ctes.columns[0])
                coluna_filial = next((c for c in df_ctes.columns if "FILIAL" in c.upper()), None)

                resultado = []
                zip_saida_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_saida_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for _, linha in df_ctes.iterrows():
                        val_bruto = str(linha[coluna_cte]).strip()
                        numero_cte = extrair_apenas_digitos_cte(val_bruto)
                        chave_pl = re.sub(r'\D', '', val_bruto)

                        # Match por número simples ou por Chave de 44 dígitos
                        xml = None
                        if numero_cte in xml_por_numero:
                            xml = xml_por_numero[numero_cte]
                        elif chave_pl in xml_por_chave:
                            xml = xml_por_chave[chave_pl]

                        if xml:
                            filial_raw = str(linha[coluna_filial]).split('.')[0].strip() if coluna_filial else "0101"
                            filial = filial_raw.zfill(4)

                            descricao = str(linha[coluna_desc]).strip() if coluna_desc else ""

                            cod_produto, tes = determinar_produto_e_tes(descricao, xml["origem"], xml["destino"], df_regras)

                            nome_pasta = f"{filial} {cod_produto} - {tes}"
                            caminho_no_zip = f"arquivos_organizados/{nome_pasta}/{xml['filename']}"

                            zip_out.writestr(caminho_no_zip, xml["content"])

                            resultado.append({
                                "CTE": numero_cte,
                                "Filial": filial,
                                "Descrição Original": descricao,
                                "Cod Produto": cod_produto,
                                "TES": tes,
                                "Nome Pasta": nome_pasta,
                                "CFOP": xml["cfop"],
                                "UF Origem": xml["origem"],
                                "UF Destino": xml["destino"],
                                "Tipo Operação": "Intraestadual" if xml["origem"] == xml["destino"] else "Interestadual",
                                "Arquivo XML": xml["filename"]
                            })

                    if resultado:
                        df_res = pd.DataFrame(resultado)
                        excel_buffer = io.BytesIO()
                        df_res.to_excel(excel_buffer, index=False)
                        zip_out.writestr("arquivos_organizados/resultado_completo.xlsx", excel_buffer.getvalue())

                # Display
                if resultado:
                    st.success(f"🎉 Sucesso! {len(resultado)} de {len(df_ctes)} registros foram organizados com sucesso!")
                    st.download_button(
                        label="📥 Baixar Arquivos Organizados (.ZIP)",
                        data=zip_saida_buffer.getvalue(),
                        file_name="CTEs_Organizados.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                else:
                    st.error("⚠️ Nenhum vínculo encontrado entre a coluna selecionada e os XMLs:")
                    st.info(f"📌 Coluna pesquisada: **'{coluna_cte}'**")
                    
                    diag_col1, diag_col2 = st.columns(2)
                    with diag_col1:
                        st.subheader("📊 Planilha (primeiros 5):")
                        for item in amostra_planilha:
                            st.write(f"- `{item}`")

                    with diag_col2:
                        st.subheader("📄 XMLs no ZIP (primeiros 5):")
                        for item in amostra_xmls:
                            st.write(f"- `{item}`")

            except Exception as e:
                st.error(f"Erro no processamento: {e}")
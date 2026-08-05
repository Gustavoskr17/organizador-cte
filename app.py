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
st.markdown("Faça o upload dos 2 arquivos abaixo para organizar os XMLs em pastas e gerar o relatório.")

# ==========================================================
# FUNÇÕES DE TRATAMENTO E DETERMINAÇÃO DO PRODUTO
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

def determinar_codigo_produto(descricao, filial="0101", cfop=""):
    """
    Determina o Código do Produto (6 dígitos) com suporte às contas 
    exatas da Filial 0202 (Solar) e regras padrão para as demais.
    """
    desc = normalizar_texto(descricao)
    cfop_str = str(cfop).strip()
    filial_clean = str(filial).split('.')[0].strip().zfill(4)

    # ==================================================================
    # REGRAS EXCLUSIVAS PARA FILIAL 0202 (SOLAR)
    # ==================================================================
    if filial_clean == "0202":
        if "IMPORTACAO" in desc or cfop_str.startswith("3"):
            if "31486" in desc:
                return "031486"
            return "048723"

        elif cfop_str in ["5910", "6910"] or any(k in desc for k in ["BRINDE", "AMOSTRA", "DOACAO", "GIFT"]):
            return "000042"

        elif any(k in desc for k in ["BONIFICACAO", "BONIF"]):
            return "000043"

        elif cfop_str in ["5915", "6915", "5949", "6949"] or any(k in desc for k in ["GARANTIA", "RMA", "REPARO", "CONSERTO", "ASSISTENCIA"]):
            return "000039"

        elif cfop_str in ["5410", "5411", "6410", "6411"] or any(k in desc for k in ["DEVOLUCAO DE VENDA", "DEV VENDA", "RETORNO DE VENDA"]):
            return "000040"

        elif cfop_str in ["5201", "5202", "6201", "6202"] or any(k in desc for k in ["DEVOLUCAO DE COMPRA", "DEV COMPRA", "RETORNO FORNECEDOR"]):
            return "000041"

        elif cfop_str in ["5151", "5152", "6151", "6152", "5357", "6357"] or any(k in desc for k in ["TRANSF", "TRANSFERENCIA"]):
            return "000037"

        elif any(k in desc for k in ["COMPRA", "COMPRAS", "INSUMO", "MATERIA PRIMA", "FORNECEDOR"]):
            return "000030"

        elif cfop_str in ["5551", "6551", "5556", "6556"] or any(k in desc for k in ["USO", "CONSUMO", "IMOBILIZADO", "ATIVO FIXO", "ESCRITORIO"]):
            return "000038"

        elif cfop_str in ["5352", "5353", "6352", "6353"] or "VENDA" in desc or ("REMESSA" in desc and "ORDEM" in desc) or "CONTA E ORDEM" in desc:
            return "028197"

        # Fallback Solar
        return "028197"

    # ==================================================================
    # DEMAIS FILIAIS (PADRÃO)
    # ==================================================================
    if any(k in desc for k in ["FUNCIONARIO", "FUNCIONARIOS", "FESTA"]):
        return "052364"

    elif "PRINCIPAL" in desc:
        return "048727"

    elif "NENHUMA" in desc or cfop_str.startswith("3") or any(k in desc for k in ["IMPORTACAO", "DESEMBARACO", "PORTUARIO"]):
        return "029975"

    elif cfop_str in ["5910", "6910"] or any(k in desc for k in ["BRINDE", "AMOSTRA", "DOACAO", "GIFT"]):
        return "051066"

    elif any(k in desc for k in ["BONIFICACAO", "BONIF"]):
        return "051068"

    elif cfop_str in ["5915", "6915", "5949", "6949"] or any(k in desc for k in ["GARANTIA", "RMA", "REPARO", "CONSERTO", "ASSISTENCIA"]):
        return "051061"

    elif cfop_str in ["5410", "5411", "6410", "6411"] or any(k in desc for k in ["DEVOLUCAO DE VENDA", "DEV VENDA", "RETORNO DE VENDA"]):
        return "051063"

    elif cfop_str in ["5201", "5202", "6201", "6202"] or any(k in desc for k in ["DEVOLUCAO DE COMPRA", "DEV COMPRA", "RETORNO FORNECEDOR"]):
        if any(k in desc for k in ["IND", "INDUSTRIA"]):
            return "051065"
        return "051064"

    elif cfop_str in ["5151", "5152", "6151", "6152", "5357", "6357"] or any(k in desc for k in ["TRANSF", "TRANSFERENCIA"]):
        if any(k in desc for k in ["IND", "INDUSTRIA"]):
            return "051057"
        return "051054"

    elif any(k in desc for k in ["COMPRA", "COMPRAS", "INSUMO", "MATERIA PRIMA", "FORNECEDOR"]):
        if any(k in desc for k in ["COM", "COMERCIO"]):
            return "051049"
        return "051047"

    elif cfop_str in ["5551", "6551", "5556", "6556"] or any(k in desc for k in ["USO", "CONSUMO", "IMOBILIZADO", "ATIVO FIXO", "ESCRITORIO"]):
        return "051060"

    elif cfop_str in ["5352", "5353", "6352", "6353"] or "VENDA" in desc or ("REMESSA" in desc and "ORDEM" in desc) or "CONTA E ORDEM" in desc:
        return "028197"

    return "028197"

# ==========================================================
# INTERFACE DO STREAMLIT (UPLOADS)
# ==========================================================

col1, col2 = st.columns(2)

with col1:
    arquivo_excel_ctes = st.file_uploader("1. Planilha de CT-es (excel.xlsx)", type=["xlsx"])

with col2:
    zip_xmls = st.file_uploader("2. Pasta compactada com XMLs (.zip)", type=["zip"])

coluna_cte_selecionada = None
if arquivo_excel_ctes:
    try:
        df_temp = pd.read_excel(arquivo_excel_ctes, nrows=5)
        df_temp.columns = df_temp.columns.str.strip()
        
        idx_padrao = 0
        for i, col in enumerate(df_temp.columns):
            if any(term in col.upper() for term in ["CTE", "NCTE", "DOC", "CHAVE", "NUMERO", "NÚMERO"]):
                idx_padrao = i
                break

        coluna_cte_selecionada = st.selectbox(
            "🎯 Selecione a coluna da planilha com o NÚMERO do CT-e ou CHAVE DE ACESSO:",
            options=list(df_temp.columns),
            index=idx_padrao
        )
    except Exception:
        pass

# ==========================================================
# BOTÃO DE PROCESSAMENTO
# ==========================================================

if st.button("🚀 Processar e Organizar CT-es", type="primary", use_container_width=True):
    if not arquivo_excel_ctes or not zip_xmls:
        st.error("⚠️ Por favor, faça o upload da Planilha e do arquivo .ZIP para continuar!")
    else:
        with st.spinner("Processando arquivos... Aguarde um instante."):
            try:
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

                df_ctes = pd.read_excel(arquivo_excel_ctes)
                df_ctes.columns = df_ctes.columns.str.strip()

                coluna_cte = coluna_cte_selecionada or df_ctes.columns[0]

                amostra_planilha = [
                    f"Valor na coluna '{coluna_cte}': {val}" 
                    for val in df_ctes[coluna_cte].dropna().head(5)
                ]

                coluna_desc = next((c for c in df_ctes.columns if "DESC" in c.upper()), df_ctes.columns[1] if len(df_ctes.columns) > 1 else df_ctes.columns[0])
                coluna_filial = next((c for c in df_ctes.columns if "FILIAL" in c.upper()), None)
                
                coluna_tes = None
                if len(df_ctes.columns) >= 48:
                    coluna_tes = df_ctes.columns[47]
                else:
                    coluna_tes = next((c for c in df_ctes.columns if "TES" in c.upper()), None)

                resultado = []
                zip_saida_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_saida_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for _, linha in df_ctes.iterrows():
                        val_bruto = str(linha[coluna_cte]).strip()
                        numero_cte = extrair_apenas_digitos_cte(val_bruto)
                        chave_pl = re.sub(r'\D', '', val_bruto)

                        xml = None
                        if numero_cte in xml_por_numero:
                            xml = xml_por_numero[numero_cte]
                        elif chave_pl in xml_por_chave:
                            xml = xml_por_chave[chave_pl]

                        if xml:
                            filial_raw = str(linha[coluna_filial]).split('.')[0].strip() if coluna_filial else "0101"
                            filial = filial_raw.zfill(4)

                            descricao = str(linha[coluna_desc]).strip() if coluna_desc else ""
                            cod_produto = determinar_codigo_produto(descricao=descricao, filial=filial, cfop=xml["cfop"])

                            tes_raw = str(linha[coluna_tes]).split('.')[0].strip() if coluna_tes and pd.notna(linha[coluna_tes]) else ""
                            tes = tes_raw.zfill(3) if tes_raw else "000"

                            nome_pasta = f"{filial} {cod_produto} - {tes}"
                            caminho_no_zip = f"arquivos_organizados/{nome_pasta}/{xml['filename']}"

                            zip_out.writestr(caminho_no_zip, xml["content"])

                            resultado.append({
                                "CTE": numero_cte,
                                "Filial": filial,
                                "Descrição Original": descricao,
                                "Cod Produto": cod_produto,
                                "TES (Planilha AV)": tes,
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

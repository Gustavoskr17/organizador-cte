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
# FUNÇÕES DE TRATAMENTO E REGRAS FIXAS
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

def determinar_produto_e_tes(descricao, uf_origem, uf_destino, filial="0101", cfop=""):
    """
    Regras de Produto (6 dígitos) e TES com base na descrição, Filial, UF e CFOP.
    """
    desc = normalizar_texto(descricao)
    uf_ori = str(uf_origem).strip().upper()
    uf_dest = str(uf_destino).strip().upper()
    filial_clean = str(filial).split('.')[0].strip().zfill(4)

    # Verifica se a operação é intraestadual (mesmo estado)
    eh_intraestadual = (uf_ori == uf_dest) and (uf_ori != "")

    # Regra da Filial 0105 / Espírito Santo (ES):
    # CT-es da filial 0105 ou ES são tributados mesmo se forem intraestaduais
    eh_tributado_como_inter = (filial_clean == "0105") or (uf_ori == "ES") or (uf_dest == "ES")

    # Define se deve aplicar a regra de Interestadual / Tributado
    usar_regra_inter = not eh_intraestadual or eh_tributado_como_inter

    # ----------------------------------------------------
    # MApeamento de Produtos (6 dígitos) e TES
    # ----------------------------------------------------

    # 1. FRETE SOBRE VENDAS / REMESSA CONTA E ORDEM
    if "VENDA" in desc or ("REMESSA" in desc and "CONTA" in desc and "ORDEM" in desc):
        return "028197", ("044" if usar_regra_inter else "045")

    # 2. FRETE DE TRANSFERÊNCIA
    elif "TRANSF" in desc or "TRANSFERENCIA" in desc:
        return "051054", ("455" if usar_regra_inter else "054")

    # 3. FRETE RMA / GARANTIA / TROCA
    elif "RMA" in desc or "GARANTIA" in desc or "TROCA" in desc:
        return "051061", ("052" if usar_regra_inter else "054")

    # 4. FRETE DEVOLUÇÃO DE VENDA
    elif "DEVOLUCAO DE VENDA" in desc or "DEV VENDA" in desc:
        return "051063", ("455" if usar_regra_inter else "480")

    # 5. FRETE DEVOLUÇÃO DE COMPRA
    elif "DEVOLUCAO" in desc or "DEV" in desc:
        return "051064", ("051" if usar_regra_inter else "480")

    # 6. FRETE BRINDES / BONIFICAÇÃO
    elif "BRINDE" in desc:
        return "051066", "356"
    elif "BONIFICACAO" in desc or "BONIF" in desc:
        return "051068", "052"

    # 7. FRETE IMPORTAÇÃO
    elif "IMPORTACAO" in desc or "IMPORT" in desc:
        return "029975", "480"

    # 8. FRETE COMPRAS INDÚSTRIA
    elif "COMPRA" in desc:
        return "051047", "454"

    # 9. FRETE USO E CONSUMO / IMOBILIZADO
    elif "USO" in desc or "CONSUMO" in desc or "IMOBILIZADO" in desc:
        return "051060", "356"

    # Padrão para não mapeados
    return "NAO_MAPEADO", "000"

# ==========================================================
# INTERFACE DO STREAMLIT (UPLOADS)
# ==========================================================

col1, col2 = st.columns(2)

with col1:
    arquivo_excel_ctes = st.file_uploader("1. Planilha de CT-es (excel.xlsx)", type=["xlsx"])

with col2:
    zip_xmls = st.file_uploader("2. Pasta compactada com XMLs (.zip)", type=["zip"])

# Seleção manual de coluna se necessário
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
                # 1. Ler todos os XMLs
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

                # 2. Ler Planilha de CT-es
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

                        # Match por número ou por Chave de Acesso
                        xml = None
                        if numero_cte in xml_por_numero:
                            xml = xml_por_numero[numero_cte]
                        elif chave_pl in xml_por_chave:
                            xml = xml_por_chave[chave_pl]

                        if xml:
                            filial_raw = str(linha[coluna_filial]).split('.')[0].strip() if coluna_filial else "0101"
                            filial = filial_raw.zfill(4)

                            descricao = str(linha[coluna_desc]).strip() if coluna_desc else ""

                            # Determina o produto (6 dígitos) e TES considerando as regras de Filial/UF/CFOP
                            cod_produto, tes = determinar_produto_e_tes(
                                descricao=descricao, 
                                uf_origem=xml["origem"], 
                                uf_destino=xml["destino"], 
                                filial=filial, 
                                cfop=xml["cfop"]
                            )

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

                # Exibição de Resultados
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

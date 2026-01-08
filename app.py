# coding: utf-8
import streamlit as st
import pandas as pd
import requests
from datetime import date
from urllib.parse import urlencode
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Download de CT-e XMLs", layout="centered")
st.title("📦 Download de CT-e (Lote XML) - MultiTransportador")

st.markdown("Envie a planilha com colunas: **Empresa, CNPJ, Usuario, Senha**")

# Session State
if "resumo_ctes" not in st.session_state:
    st.session_state.resumo_ctes = []

if "resultados" not in st.session_state:
    st.session_state.resultados = []

if "arquivos_cte" not in st.session_state:
    st.session_state.arquivos_cte = {}

if "arquivos_mdfe" not in st.session_state:
    st.session_state.arquivos_mdfe = {}

arquivo = st.file_uploader("📁 Enviar planilha", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    data_inicial = st.date_input("📅 Data Inicial", value=date(2026, 1, 1))
with col2:
    data_final = st.date_input("📅 Data Final", value=date(2026, 1, 31))

if arquivo and st.button("⬇️ Iniciar Downloads"):
    df = pd.read_excel(arquivo)

    # Limpa dados anteriores
    st.session_state.resumo_ctes.clear()
    st.session_state.resultados.clear()
    st.session_state.arquivos_cte.clear()
    st.session_state.arquivos_mdfe.clear()

    for idx, row in df.iterrows():
        empresa = row["Empresa"]
        cnpj = str(row["CNPJ"])
        usuario = str(row["Usuario"])
        senha = str(row["Senha"])

        st.write(f"🔐 Logando para empresa: **{empresa}**...")

        sessao = requests.Session()
        sessao.headers.update({
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://brf.multitransportador.com.br/Login",
            "Origin": "https://brf.multitransportador.com.br"
        })

        login_page = sessao.get("https://brf.multitransportador.com.br/Login")
        if login_page.status_code != 200:
            st.session_state.resultados.append({"Empresa": empresa, "CNPJ": cnpj, "Status": "Erro acesso"})
            continue

        payload_login = {
            "Usuario": usuario,
            "Senha": senha
        }
        resposta_login = sessao.post("https://brf.multitransportador.com.br/Login", data=payload_login, allow_redirects=False)

        cookies = sessao.cookies.get_dict()
        if "SGT.WebAdmin.Auth" not in cookies:
            st.session_state.resultados.append({"Empresa": empresa, "CNPJ": cnpj, "Status": "Login falhou"})
            continue

        st.success(f"✅ Login realizado com sucesso para: {empresa}")

        parametros = {
            "DataEmissaoInicial": data_inicial.strftime("%d/%m/%Y"),
            "DataEmissaoFinal": data_final.strftime("%d/%m/%Y"),
        }

        url_cte = (
            "https://brf.multitransportador.com.br/ConsultaCTe/DownloadLoteXML?"
            + urlencode(parametros)
        )

        url_mdfe = (
            "https://brf.multitransportador.com.br/ConsultaMDFe/DownloadLoteXML?"
            + urlencode(parametros)
        )

        # Baixa CT-e
        resposta_cte = sessao.get(url_cte)
        if resposta_cte.status_code == 200 and "application/zip" in resposta_cte.headers.get("Content-Type", ""):
            st.success(f"📥 CT-e baixado para {empresa}")
            st.session_state.arquivos_cte[cnpj] = resposta_cte.content

            zip_bytes = BytesIO(resposta_cte.content)
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                for nome_arquivo in zip_ref.namelist():
                    if nome_arquivo.lower().endswith(".xml"):
                        with zip_ref.open(nome_arquivo) as xml_file:
                            try:
                                tree = ET.parse(xml_file)
                                root = tree.getroot()
                                ns = {"ns": root.tag.split("}")[0].strip("{")}
                                ide = root.find(".//ns:ide", ns)
                                infCTe = root.find(".//ns:infCte", ns)
                                emit = root.find(".//ns:emit/ns:xNome", ns)
                                dest = root.find(".//ns:dest/ns:xNome", ns)
                                vPrest = root.find(".//ns:vPrest/ns:vRec", ns)
                                status = root.find(".//ns:protCTe/ns:infProt/ns:xMotivo", ns)

                                chave = infCTe.attrib.get("Id", "").replace("CTe", "") if infCTe is not None else ""

                                st.session_state.resumo_ctes.append({
                                    "Empresa": empresa,
                                    "CNPJ": cnpj,
                                    "Número": ide.findtext("ns:nCT", "", ns) if ide is not None else "",
                                    "Série": ide.findtext("ns:serie", "", ns) if ide is not None else "",
                                    "Chave": chave,
                                    "Data de Emissão": ide.findtext("ns:dhEmi", "", ns)[:10] if ide is not None else "",
                                    "Status": status.text if status is not None else "Autorizado",
                                    "Valor": vPrest.text if vPrest is not None else "",
                                    "Emitente": emit.text if emit is not None else "",
                                    "Destinatário": dest.text if dest is not None else "",
                                })

                            except Exception as e:
                                st.warning(f"Erro ao processar XML: {nome_arquivo} - {e}")
        else:
            st.session_state.arquivos_cte[cnpj] = None

        # Baixa MDF-e
        resposta_mdfe = sessao.get(url_mdfe)
        if resposta_mdfe.status_code == 200 and "application/zip" in resposta_mdfe.headers.get("Content-Type", ""):
            st.success(f"📥 MDF-e baixado para {empresa}")
            st.session_state.arquivos_mdfe[cnpj] = resposta_mdfe.content
        else:
            st.session_state.arquivos_mdfe[cnpj] = None

        st.session_state.resultados.append({"Empresa": empresa, "CNPJ": cnpj, "Status": "✅ Download OK"})

# 🔽 Exibe botões de download dos arquivos XML
if st.session_state.arquivos_cte:
    st.markdown("## 📦 Lotes CT-e por Empresa")
    for cnpj, zip_data in st.session_state.arquivos_cte.items():
        if zip_data:
            st.download_button(
                label=f"📥 Baixar CT-e - {cnpj}",
                data=zip_data,
                file_name=f"CTe_{cnpj}.zip",
                mime="application/zip",
                key=f"cte_{cnpj}"
            )

if st.session_state.arquivos_mdfe:
    st.markdown("## 🚛 Lotes MDF-e por Empresa")
    for cnpj, zip_data in st.session_state.arquivos_mdfe.items():
        if zip_data:
            st.download_button(
                label=f"📥 Baixar MDF-e - {cnpj}",
                data=zip_data,
                file_name=f"MDFe_{cnpj}.zip",
                mime="application/zip",
                key=f"mdfe_{cnpj}"
            )

# 🔽 Download da planilha Excel com resumo
if st.session_state.resumo_ctes:
    df_resumo = pd.DataFrame(st.session_state.resumo_ctes)
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_resumo.to_excel(writer, index=False, sheet_name="Resumo_CTes")

    st.success("✅ Planilha de resumo gerado com sucesso!")

    st.download_button(
        label="📄 Baixar Resumo CT-es",
        data=output_excel.getvalue(),
        file_name="Resumo_CTes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="resumo_excel"
    )

# 🔽 Tabela resumo das operações
if st.session_state.resultados:
    st.markdown("## 📊 Resumo das operações:")
    st.dataframe(pd.DataFrame(st.session_state.resultados))

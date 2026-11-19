import io
import os
import re
import subprocess
from datetime import datetime
import cv2
import numpy as np
import openpyxl
import streamlit as st
from playwright.sync_api import sync_playwright

# ================= CONFIGURAÇÃO DA PÁGINA =================
st.set_page_config(
    page_title="Reembolso NFC-e | Eco Brasil",
    page_icon="🧾",
    layout="centered",
)

CAMINHO_MODELO_PADRAO = "modelo_reembolso.xlsx"


# Garante que o binário do Firefox do Playwright esteja instalado no ambiente cloud
@st.cache_resource
def instalar_playwright():
    subprocess.run(["playwright", "install", "firefox"], check=False)


instalar_playwright()


# ================= TELA DE ACESSO (SENHA CORPORATIVA) =================
def verificar_acesso():
    senha_correta = st.secrets.get("SENHA_ACESSO", "ecobrasil2026")
    if st.session_state.get("autenticado"):
        return True

    st.title("🔒 Acesso Restrito — Reembolso NFC-e")
    st.markdown("Digite a senha corporativa para acessar o sistema de prestação de contas:")
    senha_digitada = st.text_input("Senha de acesso", type="password")
    if st.button("Entrar", type="primary"):
        if senha_digitada == senha_correta:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    return False


if not verificar_acesso():
    st.stop()


# ================= INTERFACE PRINCIPAL =================
st.title("🧾 Prestação de Contas Automática (NFC-e)")
st.markdown(
    "Tire foto dos cupons fiscais ou selecione da galeria. "
    "O sistema lê os QR Codes, consulta a SEFAZ-SP e gera a planilha pronta para envio."
)

with st.sidebar:
    st.header("⚙️ Dados do Colaborador")
    classe_valor = st.text_input(
        "Classe de Valor",
        value="105001 - GESTÃO CORPORATIVA",
    )
    centro_custo = st.text_input(
        "Centro de Custo",
        value="511001 - GESTAO TECNOLOGIA DA INFORMACAO",
    )
    st.divider()
    usar_planilha_propria = st.checkbox(
        "Continuar uma planilha já iniciada",
        value=False,
        help="Marque apenas se quiser adicionar notas em um arquivo .xlsx que já possui lançamentos.",
    )


# ================= FUNÇÕES DE EXTRAÇÃO E PARSER =================
def ler_qrcode_bytes(file_bytes: bytes) -> str:
    arr = np.frombuffer(file_bytes, np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError("Arquivo de imagem inválido ou corrompido.")

    detector = cv2.QRCodeDetector()
    url, _, _ = detector.detectAndDecode(img)

    if not url:
        cinza = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        for limiar in [90, 110, 130, 150, 170]:
            _, thresh = cv2.threshold(cinza, limiar, 255, cv2.THRESH_BINARY)
            url, _, _ = detector.detectAndDecode(thresh)
            if url:
                break

    if not url:
        raise ValueError("QR Code ilegível, borrado ou cortado na foto.")

    return url.strip()


def extrair_campos_texto(texto: str) -> dict:
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]

    emitente = ""
    for i, linha in enumerate(linhas):
        if "CNPJ:" in linha and i > 0:
            emitente = linhas[i - 1]
            break

    match_num = re.search(r"Número:\s*(\d+)", texto)
    numero_nf = match_num.group(1).zfill(9) if match_num else ""

    match_data = re.search(r"Emissão:\s*(\d{2}/\d{2}/\d{4})", texto)
    data_emissao = match_data.group(1) if match_data else ""

    match_valor = re.search(r"Valor a pagar R\$:\s*([\d.,]+)", texto) or re.search(
        r"Valor total R\$:\s*([\d.,]+)", texto
    )
    valor_total = 0.0
    if match_valor:
        valor_str = match_valor.group(1).replace(".", "").replace(",", ".")
        valor_total = float(valor_str)

    itens = []
    for i, linha in enumerate(linhas):
        if "(Código:" in linha and i > 0:
            candidato = linhas[i - 1]
            eh_numero = re.fullmatch(r"[\d.,]+", candidato) is not None
            eh_endereco = any(tag in candidato.upper() for tag in ["AV.", "RUA ", "ROD.", "SP", "CEP"])
            if not eh_numero and not eh_endereco:
                itens.append(candidato)

    return {
        "emitente": emitente,
        "numero_nf": numero_nf,
        "data_emissao": data_emissao,
        "valor_total": valor_total,
        "itens": itens,
    }


def consultar_sefaz_em_lote(lista_urls: list[tuple[str, str]]) -> tuple[list[dict], list[tuple[str, str]]]:
    notas_validas = []
    falhas = []

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True)
        for nome_arq, url in lista_urls:
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:132.0) Gecko/20100101 Firefox/132.0",
                locale="pt-BR",
            )
            page = context.new_page()
            try:
                page.goto(url, wait_until="networkidle", timeout=25000)
                texto_pagina = page.inner_text("body")

                if "The requested URL was rejected" in texto_pagina:
                    falhas.append((nome_arq, "Bloqueio temporário do firewall da SEFAZ. Tente reenviar esta foto."))
                    continue

                dados = extrair_campos_texto(texto_pagina)
                if not dados["numero_nf"] and dados["valor_total"] == 0.0:
                    falhas.append((nome_arq, "Portal da SEFAZ não retornou os dados da nota."))
                    continue

                dados["arquivo"] = nome_arq
                notas_validas.append(dados)

            except Exception as e:
                falhas.append((nome_arq, f"Falha ao acessar SEFAZ: {str(e)[:70]}"))
            finally:
                context.close()

        browser.close()

    return notas_validas, falhas


def classificar_despesa(emitente: str, itens: list[str]) -> str:
    conteudo = f"{emitente} {' '.join(itens)}".upper()
    if any(k in conteudo for k in ["POSTO", "COMBUST", "GASOLINA", "ETANOL", "DIESEL"]):
        return "COMBUSTÍVEL"
    if any(k in conteudo for k in ["HOTEL", "POUSADA", "HOSPEDAGEM"]):
        return "HOSPEDAGEM"
    if any(k in conteudo for k in ["UBER", "99", "TAXI", "PEDAGIO", "ESTACIONAMENTO", "PASSAGEM"]):
        return "TRANSPORTE"
    if any(k in conteudo for k in ["OFICINA", "BORRACHARIA", "GUINCHO", "AUTO PECAS"]):
        return "EMERGÊNCIA VEICULAR"
    return "ALIMENTAÇÃO"


def checar_duplicidade(ws, dados: dict, lancadas_sessao: list[dict]) -> str | None:
    for row in range(30, 71):
        nf_plan = str(ws[f"C{row}"].value or "").strip()
        obs_plan = str(ws[f"F{row}"].value or "").strip().upper()
        val_plan = ws[f"J{row}"].value
        data_plan = ws[f"B{row}"].value
        data_str = data_plan.strftime("%d/%m/%Y") if isinstance(data_plan, datetime) else str(data_plan or "").strip()

        if dados["numero_nf"] and nf_plan == dados["numero_nf"]:
            return f"Já consta na linha {row} da planilha (NF {nf_plan})."

        if (
            dados["emitente"]
            and dados["emitente"].upper() in obs_plan
            and data_str == dados["data_emissao"]
            and isinstance(val_plan, (int, float))
            and round(float(val_plan), 2) == round(dados["valor_total"], 2)
        ):
            return f"Lançamento idêntico na linha {row} ({dados['emitente']} - R$ {dados['valor_total']:.2f})."

    for outra in lancadas_sessao:
        if dados["numero_nf"] and outra["numero_nf"] == dados["numero_nf"]:
            return f"Nota repetida com a foto '{outra['arquivo']}' enviada neste mesmo lote."

    return None


def gerar_planilha_final(planilha_bytes: bytes, notas_para_inserir: list[dict], classe_val: str, centro_cus: str) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(planilha_bytes))
    ws = wb["RELATÓRIO DE CONTAS"]

    for dados in notas_para_inserir:
        linha_destino = None
        for row in range(30, 71):
            if ws[f"B{row}"].value is None and ws[f"C{row}"].value is None:
                linha_destino = row
                break

        if linha_destino is None:
            raise RuntimeError("A planilha atingiu o limite máximo (linha 70).")

        categoria = classificar_despesa(dados["emitente"], dados["itens"])
        resumo_itens = ", ".join(dados["itens"][:3])
        observacao = f"{dados['emitente']} - {resumo_itens}" if resumo_itens else dados["emitente"]
        data_dt = datetime.strptime(dados["data_emissao"], "%d/%m/%Y") if dados["data_emissao"] else ""

        ws[f"B{linha_destino}"] = data_dt
        ws[f"B{linha_destino}"].number_format = "DD/MM/YYYY"
        ws[f"C{linha_destino}"] = dados["numero_nf"]
        ws[f"D{linha_destino}"] = categoria
        ws[f"F{linha_destino}"] = observacao
        ws[f"H{linha_destino}"] = classe_val
        ws[f"I{linha_destino}"] = centro_cus
        ws[f"J{linha_destino}"] = dados["valor_total"]
        ws[f"J{linha_destino}"].number_format = "#,##0.00"

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


# ================= UPLOAD E EXECUÇÃO =================
planilha_bytes = None
nome_planilha_base = "Prestacao_de_Contas_Preenchida.xlsx"

if usar_planilha_propria:
    arq_excel = st.file_uploader("Envie sua planilha já iniciada (.xlsx)", type=["xlsx"])
    if arq_excel:
        planilha_bytes = arq_excel.getvalue()
        nome_planilha_base = f"Atualizada_{arq_excel.name}"
else:
    if os.path.exists(CAMINHO_MODELO_PADRAO):
        with open(CAMINHO_MODELO_PADRAO, "rb") as f:
            planilha_bytes = f.read()
        st.caption("📄 Usando automaticamente a planilha padrão oficial da empresa.")
    else:
        arq_excel = st.file_uploader("Envie a planilha modelo (.xlsx)", type=["xlsx"])
        if arq_excel:
            planilha_bytes = arq_excel.getvalue()

arquivos_fotos = st.file_uploader(
    "📸 Selecione ou tire foto dos cupons fiscais (QR Code)",
    type=["jpg", "jpeg", "png", "bmp", "webp"],
    accept_multiple_files=True,
)

if planilha_bytes and arquivos_fotos:
    if st.button("🚀 Processar Notas Fiscais", type="primary", use_container_width=True):
        with st.status("Lendo QR Codes e consultando a SEFAZ-SP...", expanded=True) as status:
            urls_para_consultar = []
            ilegiveis = []

            for foto in arquivos_fotos:
                try:
                    url = ler_qrcode_bytes(foto.getvalue())
                    urls_para_consultar.append((foto.name, url))
                except Exception as e:
                    ilegiveis.append((foto.name, str(e)))

            st.write(f"QR Codes identificados: **{len(urls_para_consultar)}** | Ilegíveis: **{len(ilegiveis)}**")

            notas_lidas, falhas_sefaz = consultar_sefaz_em_lote(urls_para_consultar)
            ilegiveis.extend(falhas_sefaz)
            status.update(label="Consulta concluída!", state="complete", expanded=False)

        wb_temp = openpyxl.load_workbook(io.BytesIO(planilha_bytes))
        ws_temp = wb_temp["RELATÓRIO DE CONTAS"]

        aprovadas_direto = []
        duplicadas_encontradas = []
        vistos = []

        for nota in notas_lidas:
            motivo = checar_duplicidade(ws_temp, nota, vistos)
            if motivo:
                nota["motivo_duplicidade"] = motivo
                duplicadas_encontradas.append(nota)
            else:
                aprovadas_direto.append(nota)
            vistos.append(nota)

        st.session_state["resultado"] = {
            "excel_bytes": planilha_bytes,
            "excel_nome": nome_planilha_base,
            "aprovadas": aprovadas_direto,
            "duplicadas": duplicadas_encontradas,
            "ilegiveis": ilegiveis,
        }

if "resultado" in st.session_state:
    res = st.session_state["resultado"]
    notas_finais = list(res["aprovadas"])

    if res["ilegiveis"]:
        st.error(f"⚠️ **{len(res['ilegiveis'])} nota(s) ilegível(is) ou com erro na leitura:**")
        for arq, motivo in res["ilegiveis"]:
            st.markdown(f"- **{arq}**: {motivo}")

    if res["duplicadas"]:
        st.warning(f"🔁 **{len(res['duplicadas'])} nota(s) repetida(s) detectada(s):**")
        for idx, dup in enumerate(res["duplicadas"]):
            incluir = st.checkbox(
                f"Inserir mesmo assim: **{dup['arquivo']}** — {dup['emitente']} | NF `{dup['numero_nf']}` | "
                f"Data: {dup['data_emissao']} | R$ {dup['valor_total']:.2f} *({dup['motivo_duplicidade']})*",
                key=f"dup_{idx}",
                value=False,
            )
            if incluir:
                notas_finais.append(dup)

    if notas_finais:
        st.success(f"✅ **{len(notas_finais)} nota(s)** prontas para gravação na planilha.")
        excel_pronto = gerar_planilha_final(res["excel_bytes"], notas_finais, classe_valor, centro_custo)
        st.download_button(
            label="📥 Baixar Planilha Preenchida (.xlsx)",
            data=excel_pronto,
            file_name=res["excel_nome"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("Nenhuma nota nova selecionada para inserção na planilha.")

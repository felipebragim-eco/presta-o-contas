import copy
import io
import math
import os
import re
import subprocess
from datetime import date, datetime
import cv2
import numpy as np
import openpyxl
from openpyxl.styles import Alignment
import streamlit as st
from playwright.sync_api import sync_playwright

# ================= CONFIGURAÇÃO DA PÁGINA =================
st.set_page_config(
    page_title="Prestação de Contas & Reembolso | Eco Brasil",
    page_icon="🧾",
    layout="centered",
)

CAMINHO_MODELO_PADRAO = "modelo.xlsx"

CLASSES_VALOR_OFICIAIS = [
    "105001 - GESTÃO CORPORATIVA",
    "100001 - ARAUCO SUCURI",
    "101001 - AZEVEDO",
    "102001 - BIENER",
    "103001 - BRACELL MG",
    "103002 - BRACELL SP",
    "104001 - FS FLORESTAL",
    "106001 - OFICINA CENTRAL",
    "107001 - GIRASSOL",
    "108001 - INPASA MA",
    "109001 - KLABIN CTL",
    "109002 - KLABIN FULL TREE",
    "109003 - KLABIN MODULO 13",
    "109004 - KLABIN SINISTRO",
    "109005 - KLABIN SPOT",
    "110001 - SAINT GOBAIN",
    "111001 - SUZANO ARACRUZ",
    "111002 - SUZANO MUCURI",
    "112001 - VETORIAL FAZ. ÁGUA LIMPA",
    "112002 - VETORIAL FAZ. ESTRELA DA BOTA",
    "112003 - VETORIAL FAZ. N. S. DE FÁTIMA",
    "112004 - VETORIAL FAZENDA BUBA",
    "112005 - VETORIAL SIDERURGICA",
    "112006 - VETORIAL VERDE MAR",
    "113001 - ZOPONE",
    "114001 - SYLVAMO",
    "114002 - SYLVAMO FULL TREE",
    "115001 - GREENPLAC FULL TREE",
]

CENTROS_CUSTO_OFICIAIS = [
    "511001 - GESTAO TECNOLOGIA DA INFORMACAO",
    "114001-APOIO OPERACIONAL",
    "201003-GESTÃO DE PLANEJAMENTO E CONTROLE  DE MANUTENÇÃO",
    "401001-PLANEJAMENTO E CONTROLE DE PRODUCAO",
    "402001 - SEGURANÇA, SAÚDE E MEIO AMBIENTE",
    "408001 - GERENCIA OPERACIONAL",
    "408002 - GERENCIA DE PCP PCM E SSMA",
    "409001-GESTÃO E DESENVOLVIMENTO OPERACIONAL",
    "501002 - DIRETORIA ADMINISTRATIVA",
    "501003 - DIRETORIA OPERACIONAL",
    "502001 - GESTAO COMERCIAL",
    "503001 - GESTAO FINANCEIRA",
    "504001 - GESTÃO DE PESSOAS",
    "505001 - JURIDICO",
    "506001 - GESTAO CONTROLADORIA",
    "507001 - GESTAO CONTABIL",
    "508001 - GERENCIA ADMINISTRATIVA",
    "509001 - GESTAO FISCAL TRIBUTARIA",
    "510001 - GESTAO DE COMPRAS",
    "512001 - ALMOXARIFADO MATRIZ",
    "513001 - GESTAO PLANEJ. MOBILIZ.E VIABILIZ.PESSOAS",
    "514001 - GESTAO PATRIMONIAL",
    "515001 - GESTAO DE SERVIÇOS COMPARTILHADOS",
    "516001 - GESTÃO DE PROCESSOS E MELHORIA CONTÍNUA",
    "518001 - APOIO ADMINISTRATIVO",
]


@st.cache_resource
def instalar_playwright():
    subprocess.run(["playwright", "install", "firefox"], check=False)


instalar_playwright()


# ================= FUNÇÕES DE MÁSCARA ESTILO BANCO =================
def formatar_cpf_mascara():
    """Aplica máscara 111.111.111-11 automaticamente ao digitar no campo CPF."""
    digitos = re.sub(r"\D", "", st.session_state.get("cpf_input", ""))[:11]
    if len(digitos) <= 3:
        formatado = digitos
    elif len(digitos) <= 6:
        formatado = f"{digitos[:3]}.{digitos[3:]}"
    elif len(digitos) <= 9:
        formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:]}"
    else:
        formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    st.session_state["cpf_input"] = formatado


def formatar_moeda_banco_mascara():
    """Formata o valor estilo aplicativo bancário (centavos da direita para a esquerda: R$ 1.234,56)."""
    digitos = re.sub(r"\D", "", st.session_state.get("adiantamento_input", ""))
    centavos = int(digitos) if digitos else 0
    valor_float = centavos / 100.0
    texto_br = f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    st.session_state["adiantamento_input"] = f"R$ {texto_br}"


def converter_moeda_para_float(texto_moeda: str) -> float:
    digitos = re.sub(r"\D", "", texto_moeda or "")
    return (int(digitos) / 100.0) if digitos else 0.0


# ================= TELA DE ACESSO =================
def verificar_acesso():
    senha_correta = st.secrets.get("SENHA_ACESSO", "ecobrasil2026")
    if st.session_state.get("autenticado"):
        return True

    st.title("Acesso Restrito — Eco Brasil Bioenergia")
    st.markdown("Digite a senha corporativa para acessar o sistema de prestação de contas e reembolso:")
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


# Inicializa estados das máscaras se ainda não existirem
if "cpf_input" not in st.session_state:
    st.session_state["cpf_input"] = ""
if "adiantamento_input" not in st.session_state:
    st.session_state["adiantamento_input"] = "R$ 0,00"


# ================= INTERFACE PRINCIPAL =================
st.title("Prestação de Contas & Reembolso (NFC-e)")
st.markdown(
    "Preencha os dados do relatório, envie as fotos dos cupons fiscais e baixe a planilha oficial pronta."
)

# 1. TIPO DE RELATÓRIO E DADOS DO SOLICITANTE
st.subheader("Identificação e Tipo de Relatório")
tipo_relatorio = st.radio(
    "Selecione o Tipo de Relatório:",
    options=["Relatório de Prestação de Contas", "Relatório de Reembolso"],
    horizontal=True,
)

col_sol1, col_sol2 = st.columns([2, 1])
with col_sol1:
    nome_solicitante = st.text_input("Nome do Solicitante *", placeholder="Ex: João da Silva")
with col_sol2:
    cpf_solicitante = st.text_input(
        "CPF * (digite apenas números)",
        key="cpf_input",
        on_change=formatar_cpf_mascara,
        placeholder="111.111.111-11",
        max_chars=14,
    )

col_resp1, col_resp2 = st.columns([2, 1])
with col_resp1:
    responsavel_area = st.text_input("Responsável da Área *", placeholder="Nome do gestor responsável")
with col_resp2:
    if tipo_relatorio == "Relatório de Prestação de Contas":
        adiantamento_texto = st.text_input(
            "Valor do Adiantamento (digite só números)",
            key="adiantamento_input",
            on_change=formatar_moeda_banco_mascara,
            help="Digite apenas os números (ex: 15000 vira R$ 150,00).",
        )
        valor_adiantamento = converter_moeda_para_float(adiantamento_texto)
        data_adiantamento = st.date_input("Data do Adiantamento", value=date.today(), format="DD/MM/YYYY")
    else:
        valor_adiantamento = 0.0
        data_adiantamento = None
        st.caption("No Relatório de Reembolso não há valor de adiantamento.")

# 2. DADOS BANCÁRIOS (PIX OU TRANSFERÊNCIA)
st.subheader("Dados Bancários para Pagamento")
forma_pagamento = st.radio(
    "Forma de Pagamento:",
    options=["PAGAMENTO VIA PIX", "TRANSFERÊNCIA BANCÁRIA"],
    horizontal=True,
)

dados_bancarios = {"forma": forma_pagamento}

if forma_pagamento == "PAGAMENTO VIA PIX":
    c_pix1, c_pix2, c_pix3 = st.columns([1, 1.5, 1.5])
    with c_pix1:
        dados_bancarios["tipo_chave"] = st.selectbox(
            "Tipo da Chave PIX *",
            options=["CPF", "TELEFONE", "E-MAIL", "CHAVE ALEATÓRIA"],
        )
    with c_pix2:
        dados_bancarios["chave_pix"] = st.text_input("Chave PIX *", placeholder="Informe a chave PIX")
    with c_pix3:
        dados_bancarios["responsavel_conta"] = st.text_input(
            "Responsável pela Conta *",
            value=nome_solicitante,
            placeholder="Titular da conta PIX",
        )
else:
    c_ted1, c_ted2, c_ted3, c_ted4 = st.columns([1.2, 1, 0.8, 1.5])
    with c_ted1:
        dados_bancarios["banco"] = st.text_input("Banco *", placeholder="Ex: Banco do Brasil / Itaú")
    with c_ted2:
        dados_bancarios["conta"] = st.text_input("Conta (com dígito) *", placeholder="Ex: 12345-6")
    with c_ted3:
        dados_bancarios["agencia"] = st.text_input("Agência *", placeholder="Ex: 0001")
    with c_ted4:
        dados_bancarios["responsavel_conta"] = st.text_input(
            "Responsável pela Conta *",
            value=nome_solicitante,
            placeholder="Titular da conta bancária",
        )

# 3. CLASSIFICAÇÃO CONTÁBIL
with st.expander("Classe de Valor e Centro de Custo (Padrão TI — clique para alterar)", expanded=False):
    col_cc1, col_cc2 = st.columns(2)
    with col_cc1:
        classe_valor = st.selectbox("Classe de Valor", options=CLASSES_VALOR_OFICIAIS, index=0)
    with col_cc2:
        centro_custo = st.selectbox("Centro de Custo", options=CENTROS_CUSTO_OFICIAIS, index=0)

    usar_planilha_propria = st.checkbox(
        "Continuar a partir de uma planilha já iniciada (.xlsx)",
        value=False,
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


def aplicar_alinhamento_sem_vazar(cell, horizontal: str = "center", wrap: bool = True, shrink: bool = False):
    """Mantém o alinhamento da célula e ativa quebra automática para não vazar das bordas."""
    alinhamento_atual = copy.copy(cell.alignment) if cell.alignment else Alignment()
    alinhamento_atual.horizontal = horizontal
    alinhamento_atual.vertical = "center"
    alinhamento_atual.wrap_text = wrap
    alinhamento_atual.shrink_to_fit = shrink
    cell.alignment = alinhamento_atual


def gerar_planilha_final(
    planilha_bytes: bytes,
    notas_para_inserir: list[dict],
    cabecalho: dict,
    bancario: dict,
    classe_val: str,
    centro_cus: str,
) -> bytes:
    """Preenche a planilha preservando 100% da estrutura, fórmulas, validações e estilos originais."""
    wb = openpyxl.load_workbook(io.BytesIO(planilha_bytes))
    ws = wb["RELATÓRIO DE CONTAS"]

    # 1. Tipo de Relatório via tabela auxiliar (V3: 1 = Reembolso | 2 = Prestação de Contas)
    ws["V3"] = 1 if cabecalho["tipo_relatorio"] == "Relatório de Reembolso" else 2

    # 2. Data de Emissão (J10)
    ws["J10"] = datetime.combine(date.today(), datetime.min.time())
    ws["J10"].number_format = "DD/MM/YYYY"

    # 3. Nome do Solicitante (B14:H14), CPF formatado com pontos e traço (I14:J14) e Responsável (B17:J17)
    ws["B14"] = cabecalho["nome_solicitante"].strip().upper()
    aplicar_alinhamento_sem_vazar(ws["B14"], horizontal="left", wrap=True)

    cpf_digitos = re.sub(r"\D", "", cabecalho["cpf"])
    if len(cpf_digitos) == 11:
        cpf_formatado = f"{cpf_digitos[:3]}.{cpf_digitos[3:6]}.{cpf_digitos[6:9]}-{cpf_digitos[9:]}"
    else:
        cpf_formatado = cabecalho["cpf"].strip()
    ws["I14"] = cpf_formatado
    ws["I14"].number_format = "@"
    aplicar_alinhamento_sem_vazar(ws["I14"], horizontal="center", wrap=False, shrink=True)

    ws["B17"] = cabecalho["responsavel_area"].strip().upper()
    aplicar_alinhamento_sem_vazar(ws["B17"], horizontal="left", wrap=True)

    # 4. Adiantamento (D19 e J20) quando for Prestação de Contas
    if cabecalho["tipo_relatorio"] == "Relatório de Prestação de Contas":
        ws["D19"] = float(cabecalho["valor_adiantamento"])
        ws["D19"].number_format = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-'
        if cabecalho["data_adiantamento"]:
            ws["J20"] = datetime.combine(cabecalho["data_adiantamento"], datetime.min.time())
            ws["J20"].number_format = "DD/MM/YYYY"
    else:
        ws["D19"] = None
        ws["J20"] = None

    # 5. Checkboxes e Dados Bancários (Linhas 24 e 27)
    if bancario["forma"] == "PAGAMENTO VIA PIX":
        ws["B24"] = True
        ws["B27"] = False

        ws["E24"] = bancario["tipo_chave"]
        aplicar_alinhamento_sem_vazar(ws["E24"], horizontal="center", wrap=True)

        ws["F24"] = bancario["chave_pix"].strip()
        aplicar_alinhamento_sem_vazar(ws["F24"], horizontal="left", wrap=True)

        ws["H24"] = bancario["responsavel_conta"].strip().upper()
        aplicar_alinhamento_sem_vazar(ws["H24"], horizontal="left", wrap=True)

        for col_limpar in ["E27", "F27", "G27", "H27"]:
            ws[col_limpar] = None
    else:
        ws["B24"] = False
        ws["B27"] = True

        ws["E27"] = bancario["banco"].strip().upper()
        aplicar_alinhamento_sem_vazar(ws["E27"], horizontal="center", wrap=True)

        ws["F27"] = bancario["conta"].strip()
        aplicar_alinhamento_sem_vazar(ws["F27"], horizontal="center", wrap=True)

        ws["G27"] = bancario["agencia"].strip()
        aplicar_alinhamento_sem_vazar(ws["G27"], horizontal="center", wrap=True)

        ws["H27"] = bancario["responsavel_conta"].strip().upper()
        aplicar_alinhamento_sem_vazar(ws["H27"], horizontal="left", wrap=True)

        for col_limpar in ["E24", "F24", "H24"]:
            ws[col_limpar] = None

    # 6. Lançamento dos Cupons / Notas Fiscais (Linhas 30 a 70)
    for dados in notas_para_inserir:
        linha_destino = None
        for row in range(30, 71):
            if ws[f"B{row}"].value is None and ws[f"C{row}"].value is None:
                linha_destino = row
                break

        if linha_destino is None:
            raise RuntimeError("A planilha atingiu o limite máximo de 41 lançamentos (linhas 30 a 70).")

        categoria = classificar_despesa(dados["emitente"], dados["itens"])
        resumo_itens = ", ".join(dados["itens"][:3])
        observacao = f"{dados['emitente']} - {resumo_itens}" if resumo_itens else dados["emitente"]
        data_dt = datetime.strptime(dados["data_emissao"], "%d/%m/%Y") if dados["data_emissao"] else ""

        ws[f"B{linha_destino}"] = data_dt
        ws[f"B{linha_destino}"].number_format = "DD/MM/YYYY"
        aplicar_alinhamento_sem_vazar(ws[f"B{linha_destino}"], horizontal="center", wrap=False)

        ws[f"C{linha_destino}"] = dados["numero_nf"]
        ws[f"C{linha_destino}"].number_format = "@"
        aplicar_alinhamento_sem_vazar(ws[f"C{linha_destino}"], horizontal="center", wrap=False)

        ws[f"D{linha_destino}"] = categoria
        aplicar_alinhamento_sem_vazar(ws[f"D{linha_destino}"], horizontal="center", wrap=True)

        ws[f"F{linha_destino}"] = observacao
        aplicar_alinhamento_sem_vazar(ws[f"F{linha_destino}"], horizontal="left", wrap=True)

        ws[f"H{linha_destino}"] = classe_val
        ws[f"H{linha_destino}"].number_format = "General"
        aplicar_alinhamento_sem_vazar(ws[f"H{linha_destino}"], horizontal="center", wrap=True)

        ws[f"I{linha_destino}"] = centro_cus
        ws[f"I{linha_destino}"].number_format = "General"
        aplicar_alinhamento_sem_vazar(ws[f"I{linha_destino}"], horizontal="center", wrap=True)

        ws[f"J{linha_destino}"] = dados["valor_total"]
        aplicar_alinhamento_sem_vazar(ws[f"J{linha_destino}"], horizontal="right", wrap=False)

        linhas_obs = math.ceil(len(observacao) / 36) if observacao else 1
        linhas_cc = math.ceil(len(centro_cus) / 24) if centro_cus else 1
        qtd_linhas_texto = max(linhas_obs, linhas_cc, 1)
        ws.row_dimensions[linha_destino].height = max(21.0, qtd_linhas_texto * 15.5)

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


# ================= UPLOAD E EXECUÇÃO =================
st.subheader("Notas Fiscais (QR Code)")

planilha_bytes = None
nome_arquivo_saida = f"{tipo_relatorio.replace(' ', '_')}.xlsx"

if usar_planilha_propria:
    arq_excel = st.file_uploader("Envie sua planilha já iniciada (.xlsx)", type=["xlsx"])
    if arq_excel:
        planilha_bytes = arq_excel.getvalue()
        nome_arquivo_saida = f"Atualizada_{arq_excel.name}"
else:
    if os.path.exists(CAMINHO_MODELO_PADRAO):
        with open(CAMINHO_MODELO_PADRAO, "rb") as f:
            planilha_bytes = f.read()
    else:
        arq_excel = st.file_uploader("Envie o arquivo modelo.xlsx", type=["xlsx"])
        if arq_excel:
            planilha_bytes = arq_excel.getvalue()

arquivos_fotos = st.file_uploader(
    "Selecione ou tire foto dos cupons fiscais (QR Code)",
    type=["jpg", "jpeg", "png", "bmp", "webp"],
    accept_multiple_files=True,
)


def validar_campos_obrigatorios() -> list[str]:
    faltando = []
    if not nome_solicitante.strip():
        faltando.append("Nome do Solicitante")
    cpf_limpo = re.sub(r"\D", "", cpf_solicitante)
    if len(cpf_limpo) != 11:
        faltando.append("CPF válido (11 dígitos)")
    if not responsavel_area.strip():
        faltando.append("Responsável da Área")
    if forma_pagamento == "PAGAMENTO VIA PIX":
        if not dados_bancarios.get("chave_pix", "").strip():
            faltando.append("Chave PIX")
        if not dados_bancarios.get("responsavel_conta", "").strip():
            faltando.append("Responsável pela Conta (PIX)")
    else:
        if not dados_bancarios.get("banco", "").strip():
            faltando.append("Banco")
        if not dados_bancarios.get("conta", "").strip():
            faltando.append("Conta")
        if not dados_bancarios.get("agencia", "").strip():
            faltando.append("Agência")
        if not dados_bancarios.get("responsavel_conta", "").strip():
            faltando.append("Responsável pela Conta (Transferência)")
    return faltando


if planilha_bytes and arquivos_fotos:
    if st.button("Processar Notas Fiscais", type="primary", use_container_width=True):
        campos_faltantes = validar_campos_obrigatorios()
        if campos_faltantes:
            st.warning(f"Preencha corretamente antes de continuar: **{', '.join(campos_faltantes)}**.")
        else:
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
                "excel_nome": nome_arquivo_saida,
                "aprovadas": aprovadas_direto,
                "duplicadas": duplicadas_encontradas,
                "ilegiveis": ilegiveis,
            }

if "resultado" in st.session_state:
    res = st.session_state["resultado"]
    notas_finais = list(res["aprovadas"])

    if res["ilegiveis"]:
        st.error(f" **{len(res['ilegiveis'])} nota(s) ilegível(is) ou com erro na leitura:**")
        for arq, motivo in res["ilegiveis"]:
            st.markdown(f"- **{arq}**: {motivo}")

    if res["duplicadas"]:
        st.warning(f" **{len(res['duplicadas'])} nota(s) repetida(s) detectada(s):**")
        for idx, dup in enumerate(res["duplicadas"]):
            incluir = st.checkbox(
                f"Inserir mesmo assim: **{dup['arquivo']}** — {dup['emitente']} | NF `{dup['numero_nf']}` | "
                f"Data: {dup['data_emissao']} | R$ {dup['valor_total']:.2f} *({dup['motivo_duplicidade']})*",
                key=f"dup_{idx}",
                value=False,
            )
            if incluir:
                notas_finais.append(dup)

    cabecalho_atual = {
        "tipo_relatorio": tipo_relatorio,
        "nome_solicitante": nome_solicitante,
        "cpf": cpf_solicitante,
        "responsavel_area": responsavel_area,
        "valor_adiantamento": valor_adiantamento,
        "data_adiantamento": data_adiantamento,
    }

    excel_pronto = gerar_planilha_final(
        res["excel_bytes"],
        notas_finais,
        cabecalho_atual,
        dados_bancarios,
        classe_valor,
        centro_custo,
    )

    st.success(f"Planilha pronta com **{len(notas_finais)} nota(s)** lançada(s) e cabeçalho preenchido.")
    st.download_button(
        label="Baixar Planilha Oficial Preenchida (.xlsx)",
        data=excel_pronto,
        file_name=res["excel_nome"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

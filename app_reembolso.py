import ctypes
import os
import re
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from tkinter import Button, Label, Radiobutton, StringVar, Tk, filedialog, messagebox
import cv2
import numpy as np
import openpyxl
import pyautogui
import pyperclip

# ================= CONFIGURAÇÕES =================
CLASSE_VALOR_PADRAO = "105001 - GESTÃO CORPORATIVA"
CENTRO_CUSTO_PADRAO = "511001 - GESTAO TECNOLOGIA DA INFORMACAO"

TEMPO_ESPERA_ABRIR_NAVEGADOR = 6.5  # Segundos para o navegador inicializar o perfil limpo
TEMPO_ESPERA_CARREGAR_SEFAZ = 7.0   # Segundos para a página da SEFAZ carregar

EXTENSOES_IMAGEM = (".jpg", ".jpeg", ".png", ".bmp", ".webp")

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.4
# =================================================


def localizar_navegadores_disponiveis() -> dict:
    """Verifica quais navegadores estão instalados na máquina."""
    candidatos = {
        "Mozilla Firefox": [
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Mozilla Firefox\firefox.exe"),
        ],
        "Google Chrome": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ],
        "Microsoft Edge": [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ],
        "Opera GX": [
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Opera GX\opera.exe"),
            r"C:\Program Files\Opera GX\opera.exe",
            r"C:\Program Files (x86)\Opera GX\opera.exe",
        ],
    }
    comandos = {
        "Mozilla Firefox": "firefox",
        "Google Chrome": "chrome",
        "Microsoft Edge": "msedge",
        "Opera GX": "opera",
    }

    disponiveis = {}
    for nome, caminhos in candidatos.items():
        exe = next((p for p in caminhos if os.path.exists(p)), shutil.which(comandos[nome]))
        if exe:
            disponiveis[nome] = exe
    return disponiveis


def escolher_navegador_gui(disponiveis: dict) -> tuple[str, str]:
    """Exibe janela para o usuário escolher o navegador com guia anônima desbloqueada."""
    if not disponiveis:
        raise FileNotFoundError("Nenhum navegador compatível (Firefox, Chrome, Edge ou Opera GX) foi encontrado.")

    selecionado = {"valor": None}

    janela = Tk()
    janela.title("Seleção de Navegador - NFC-e")
    janela.attributes("-topmost", True)
    janela.resizable(False, False)

    # Centraliza a janela na tela
    largura_win, altura_win = 470, 310
    x = (janela.winfo_screenwidth() // 2) - (largura_win // 2)
    y = (janela.winfo_screenheight() // 2) - (altura_win // 2)
    janela.geometry(f"{largura_win}x{altura_win}+{x}+{y}")

    Label(
        janela,
        text="Escolha o navegador para consultar as notas:",
        font=("Segoe UI", 11, "bold"),
    ).pack(pady=(15, 6))

    Label(
        janela,
        text="AVISO IMPORTANTE:\n"
             "O script utiliza a GUIA ANÔNIMA / PRIVATIVA para evitar bloqueios\n"
             "de cookies da SEFAZ. Escolha um navegador onde essa função esteja\n"
             "DESBLOQUEADA no seu computador.",
        font=("Segoe UI", 9),
        fg="#8a4b00",
        bg="#fff8e6",
        relief="solid",
        bd=1,
        padx=10,
        pady=8,
        justify="left",
    ).pack(padx=20, pady=(0, 12), fill="x")

    nomes = list(disponiveis.keys())
    var_opcao = StringVar(value=nomes[0])

    for nome in nomes:
        Radiobutton(
            janela,
            text=nome,
            variable=var_opcao,
            value=nome,
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(padx=40, anchor="w")

    def confirmar():
        selecionado["valor"] = var_opcao.get()
        janela.destroy()

    Button(
        janela,
        text="Continuar",
        command=confirmar,
        font=("Segoe UI", 10, "bold"),
        bg="#0066cc",
        fg="white",
        padx=20,
        pady=5,
    ).pack(pady=16)

    janela.mainloop()

    if not selecionado["valor"]:
        raise SystemExit("Nenhum navegador selecionado. Encerrando.")

    nome_escolhido = selecionado["valor"]
    return nome_escolhido, disponiveis[nome_escolhido]


def selecionar_origem_e_planilha():
    """Permite escolher uma PASTA inteira de notas ou arquivos avulsos, e depois a planilha."""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    usar_pasta = messagebox.askyesno(
        title="Modo de Seleção",
        message="Deseja selecionar uma PASTA inteira com várias notas fiscais?\n\n"
                "• Sim = Selecionar uma pasta completa\n"
                "• Não = Selecionar arquivos individuais",
    )

    imagens = []
    if usar_pasta:
        print("Selecione a PASTA contendo as imagens das notas fiscais...")
        pasta = filedialog.askdirectory(title="Selecione a Pasta com as Notas Fiscais")
        if not pasta:
            raise SystemExit("Nenhuma pasta selecionada. Encerrando.")

        for nome_arq in sorted(os.listdir(pasta)):
            if nome_arq.lower().endswith(EXTENSOES_IMAGEM):
                imagens.append(os.path.join(pasta, nome_arq))

        if not imagens:
            raise SystemExit(f"Nenhuma imagem encontrada na pasta: {pasta}")
    else:
        print("Selecione a(s) imagem(ns) do QR Code das notas fiscais...")
        arquivos = filedialog.askopenfilenames(
            title="Selecione as imagens das Notas Fiscais",
            filetypes=[("Imagens", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Todos os arquivos", "*.*")],
        )
        if not arquivos:
            raise SystemExit("Nenhuma imagem selecionada. Encerrando.")
        imagens = list(arquivos)

    print("Selecione a planilha Excel de Prestação de Contas...")
    planilha = filedialog.askopenfilename(
        title="Selecione a Planilha Excel",
        filetypes=[("Planilhas Excel", "*.xlsx")],
    )
    if not planilha:
        raise SystemExit("Nenhuma planilha selecionada. Encerrando.")

    root.destroy()
    return imagens, planilha


def ler_qrcode(caminho_imagem: str) -> str:
    """Lê a URL exata do QR Code sem alterar o hash de assinatura."""
    dados_bytes = np.fromfile(caminho_imagem, dtype=np.uint8)
    img = cv2.imdecode(dados_bytes, cv2.IMREAD_COLOR)
    if img is None:
        raise FileNotFoundError(f"Não foi possível abrir o arquivo de imagem.")

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
        raise ValueError("QR Code ilegível, borrado ou não encontrado na foto.")

    return url.strip()


def preparar_perfil_firefox(pasta_perfil: str):
    """Desativa telas de boas-vindas e popups iniciais no Firefox."""
    prefs = """
user_pref("browser.startup.homepage_override.mstone", "ignore");
user_pref("startup.homepage_welcome_url", "");
user_pref("startup.homepage_welcome_url.additional", "");
user_pref("startup.homepage_override_url", "");
user_pref("browser.aboutwelcome.enabled", false);
user_pref("datareporting.policy.firstRunURL", "");
user_pref("datareporting.policy.dataSubmissionPolicyBypassNotification", true);
user_pref("browser.shell.checkDefaultBrowser", false);
user_pref("browser.messaging-system.whatsNewPanel.enabled", false);
user_pref("trailhead.firstrun.didSeeAboutWelcome", true);
"""
    with open(os.path.join(pasta_perfil, "user.js"), "w", encoding="utf-8") as f:
        f.write(prefs.strip())


def aguardar_janela_ativa(palavras_chave: list[str], timeout: float = 15.0) -> bool:
    """Aguarda até que uma janela do navegador escolhido fique em primeiro plano."""
    user32 = ctypes.windll.user32
    inicio = time.time()
    while time.time() - inicio < timeout:
        hwnd = user32.GetForegroundWindow()
        if hwnd:
            tamanho = user32.GetWindowTextLengthW(hwnd)
            buffer = ctypes.create_unicode_buffer(tamanho + 1)
            user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
            titulo = buffer.value.lower()
            if any(p.lower() in titulo for p in palavras_chave):
                return True
        time.sleep(0.5)
    return False


def capturar_dados_no_navegador(nome_nav: str, exe_nav: str, url: str) -> dict:
    """Abre o navegador escolhido em guia anônima/privativa com perfil limpo e extrai os dados."""
    pasta_perfil_temp = tempfile.mkdtemp(prefix="nfce_nav_")

    try:
        if nome_nav == "Mozilla Firefox":
            preparar_perfil_firefox(pasta_perfil_temp)
            cmd = [exe_nav, "-no-remote", "-profile", pasta_perfil_temp, "-private-window", "about:blank"]
            palavras = ["Firefox", "Mozilla", "Navegação privativa", "Private Browsing"]
        elif nome_nav == "Google Chrome":
            cmd = [exe_nav, "--incognito", f"--user-data-dir={pasta_perfil_temp}", "--no-first-run", "--no-default-browser-check", "--start-maximized", "about:blank"]
            palavras = ["Chrome", "anônima", "Incognito"]
        elif nome_nav == "Microsoft Edge":
            cmd = [exe_nav, "--inprivate", f"--user-data-dir={pasta_perfil_temp}", "--no-first-run", "--no-default-browser-check", "--start-maximized", "about:blank"]
            palavras = ["Edge", "InPrivate"]
        else:  # Opera GX
            cmd = [exe_nav, "--private", f"--user-data-dir={pasta_perfil_temp}", "--no-first-run", "--start-maximized", "about:blank"]
            palavras = ["Opera", "Privada", "Private"]

        subprocess.Popen(cmd)
        print(f"   Aguardando a guia anônima do {nome_nav} abrir...")
        aguardar_janela_ativa(palavras, timeout=15.0)
        time.sleep(TEMPO_ESPERA_ABRIR_NAVEGADOR)

        # Garante fechamento de qualquer popup residual e foca a janela
        pyautogui.press("enter")
        time.sleep(0.4)
        largura, altura = pyautogui.size()
        pyautogui.click(largura // 2, altura // 2)
        time.sleep(0.4)

        # Cola a URL na barra de endereços
        pyperclip.copy(url)
        pyautogui.hotkey("ctrl", "l")
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.4)
        pyautogui.press("enter")

        print(f"   Aguardando {TEMPO_ESPERA_CARREGAR_SEFAZ}s para a página da SEFAZ carregar...")
        time.sleep(TEMPO_ESPERA_CARREGAR_SEFAZ)

        pyautogui.click(largura // 2, altura // 2)
        time.sleep(0.4)
        pyperclip.copy("")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)
        texto_pagina = pyperclip.paste()

        pyautogui.hotkey("alt", "f4")
        time.sleep(1.5)

        if "The requested URL was rejected" in texto_pagina:
            raise RuntimeError("A URL foi rejeitada pelo firewall da SEFAZ (tente outro navegador).")

        dados = extrair_campos_texto(texto_pagina)
        if not dados["numero_nf"] and dados["valor_total"] == 0.0:
            raise RuntimeError("Não foi possível ler os dados na página (página em branco, guia anônima bloqueada ou lentidão).")

        return dados

    finally:
        shutil.rmtree(pasta_perfil_temp, ignore_errors=True)


def extrair_campos_texto(texto: str) -> dict:
    """Extrai Emitente, Número, Data, Valor Total e apenas os produtos reais da SEFAZ-SP."""
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


def classificar_despesa(emitente: str, itens: list[str]) -> str:
    """Classifica a despesa dentro das 5 opções válidas da planilha."""
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


def verificar_duplicidade_na_planilha(ws, dados: dict) -> str | None:
    """Verifica se a nota já existe na planilha (por Número da NF ou Data + Emitente + Valor)."""
    for row in range(30, 71):
        nf_planilha = str(ws[f"C{row}"].value or "").strip()
        obs_planilha = str(ws[f"F{row}"].value or "").strip().upper()
        val_planilha = ws[f"J{row}"].value
        data_planilha = ws[f"B{row}"].value

        data_str = ""
        if isinstance(data_planilha, datetime):
            data_str = data_planilha.strftime("%d/%m/%Y")
        elif data_planilha:
            data_str = str(data_planilha).strip()

        # 1. Mesmo número de NF-e
        if dados["numero_nf"] and nf_planilha == dados["numero_nf"]:
            return f"Linha {row} já possui a NF nº {nf_planilha} ({obs_planilha})."

        # 2. Mesma Data, mesmo Emitente e mesmo Valor (caso tenha sido lançada sem o zero à esquerda)
        if (
            dados["emitente"]
            and dados["emitente"].upper() in obs_planilha
            and data_str == dados["data_emissao"]
            and isinstance(val_planilha, (int, float))
            and round(float(val_planilha), 2) == round(dados["valor_total"], 2)
        ):
            return f"Linha {row} já possui lançamento idêntico ({dados['emitente']} em {data_str} no valor de R$ {dados['valor_total']:.2f})."

    return None


def perguntar_inserir_duplicada(nome_arquivo: str, dados: dict, motivo_duplicidade: str) -> bool:
    """Mostra um alerta visual perguntando se o usuário deseja inserir a nota repetida mesmo assim."""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    resposta = messagebox.askyesno(
        title="Aviso: Nota Fiscal Repetida Detectada",
        message=(
            f"A imagem '{nome_arquivo}' parece conter uma nota repetida!\n\n"
            f"Motivo: {motivo_duplicidade}\n\n"
            f"Dados lidos agora:\n"
            f"• Emitente: {dados['emitente']}\n"
            f"• Número NF: {dados['numero_nf']}\n"
            f"• Data: {dados['data_emissao']}\n"
            f"• Valor: R$ {dados['valor_total']:.2f}\n\n"
            f"Deseja inserir esta nota na planilha MESMO ASSIM?"
        ),
        icon="warning",
    )
    root.destroy()
    return resposta


def preencher_planilha(caminho_excel: str, dados: dict, nome_arquivo: str) -> bool:
    """Insere os dados na próxima linha vazia (entre 30 e 70). Retorna True se salvou, False se pulou."""
    wb = openpyxl.load_workbook(caminho_excel)
    ws = wb["RELATÓRIO DE CONTAS"]

    motivo_dup = verificar_duplicidade_na_planilha(ws, dados)
    if motivo_dup:
        print(f"   [Aviso Duplicidade] {motivo_dup}")
        if not perguntar_inserir_duplicada(nome_arquivo, dados, motivo_dup):
            print("   -> Usuário optou por NÃO inserir a nota repetida. Pulando...")
            return False
        print("   -> Usuário optou por inserir mesmo assim.")

    linha_destino = None
    for row in range(30, 71):
        if ws[f"B{row}"].value is None and ws[f"C{row}"].value is None:
            linha_destino = row
            break

    if linha_destino is None:
        raise RuntimeError("Não há mais linhas disponíveis na tabela (linhas 30 a 70).")

    categoria = classificar_despesa(dados["emitente"], dados["itens"])
    resumo_itens = ", ".join(dados["itens"][:3])
    observacao = f"{dados['emitente']} - {resumo_itens}" if resumo_itens else dados["emitente"]

    data_dt = datetime.strptime(dados["data_emissao"], "%d/%m/%Y") if dados["data_emissao"] else ""

    ws[f"B{linha_destino}"] = data_dt
    ws[f"B{linha_destino}"].number_format = "DD/MM/YYYY"
    ws[f"C{linha_destino}"] = dados["numero_nf"]
    ws[f"D{linha_destino}"] = categoria
    ws[f"F{linha_destino}"] = observacao
    ws[f"H{linha_destino}"] = CLASSE_VALOR_PADRAO
    ws[f"I{linha_destino}"] = CENTRO_CUSTO_PADRAO
    ws[f"J{linha_destino}"] = dados["valor_total"]
    ws[f"J{linha_destino}"].number_format = "#,##0.00"

    while True:
        try:
            wb.save(caminho_excel)
            break
        except PermissionError:
            input(
                f"\n[AVISO] A planilha '{os.path.basename(caminho_excel)}' está aberta no Excel.\n"
                "Feche o arquivo no Excel e pressione ENTER aqui para salvar..."
            )

    print(f"   -> Planilha atualizada na linha {linha_destino}!")
    return True


def exibir_resumo_final(total: int, sucessos: int, puladas: int, ilegiveis: list[tuple[str, str]]):
    """Exibe resumo no console e abre janela de alerta caso alguma nota esteja ilegível."""
    print("\n" + "=" * 55)
    print(f"PROCESSAMENTO FINALIZADO:")
    print(f" • Notas lançadas com sucesso: {sucessos}/{total}")
    if puladas:
        print(f" • Notas repetidas ignoradas:  {puladas}")
    if ilegiveis:
        print(f" • Notas ilegíveis / com erro: {len(ilegiveis)}")
        for arq, motivo in ilegiveis:
            print(f"    - {arq}: {motivo}")
    print("=" * 55)

    if ilegiveis:
        lista_formatada = "\n".join([f"• {arq} ({motivo})" for arq, motivo in ilegiveis])
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showwarning(
            title="Atenção: Notas Fiscais Ilegíveis",
            message=(
                f"O processamento foi concluído ({sucessos} lançadas), porém "
                f"{len(ilegiveis)} nota(s) não puderam ser lidas e precisam ser verificadas manualmente:\n\n"
                f"{lista_formatada}"
            ),
        )
        root.destroy()


if __name__ == "__main__":
    navegadores = localizar_navegadores_disponiveis()
    nome_navegador, exe_navegador = escolher_navegador_gui(navegadores)
    print(f"Navegador selecionado: {nome_navegador}")

    lista_imagens, caminho_planilha = selecionar_origem_e_planilha()
    print(f"\nTotal de imagens encontradas para processar: {len(lista_imagens)}")

    sucessos = 0
    puladas = 0
    notas_ilegiveis = []

    for idx, img_path in enumerate(lista_imagens, start=1):
        nome_arquivo = os.path.basename(img_path)
        print(f"\n[{idx}/{len(lista_imagens)}] Processando: {nome_arquivo}")
        try:
            url_sefaz = ler_qrcode(img_path)
            print("   URL lida:", url_sefaz)

            dados_nota = capturar_dados_no_navegador(nome_navegador, exe_navegador, url_sefaz)
            print("   Dados extraídos:", dados_nota)

            inserida = preencher_planilha(caminho_planilha, dados_nota, nome_arquivo)
            if inserida:
                sucessos += 1
            else:
                puladas += 1

        except Exception as erro:
            print(f"   [ILEGÍVEL / ERRO] '{nome_arquivo}': {erro}")
            notas_ilegiveis.append((nome_arquivo, str(erro)))

    exibir_resumo_final(len(lista_imagens), sucessos, puladas, notas_ilegiveis)
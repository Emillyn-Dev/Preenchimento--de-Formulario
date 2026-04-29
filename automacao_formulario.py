#Para rodar o código, preciso estar na pasta correta (escrevo cd .\Formulário)
#Depois ativo o venv (escrevo  .\.venv\Scripts\Activate.ps1 )
#Ai coloco pra rodar


import os
import re
import time
import unicodedata

import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# =========================
# CONFIGURAÇÃO
# =========================
os.chdir(r"C:\Users\emill\Desktop\Programação\Rede_Dor\Formulário")

URL = "https://rdslprod.service-now.com/snow?id=sc_cat_item&sys_id=8d9d72d71b4ea11060d9426fe54bcbe0&referrer=popular_items"
PLANILHA = "automacao.xlsx"
SHEET = "Planilha1"

# =========================
# VERIFICA SE PLANILHA ESTÁ ABERTA
# =========================
try:
    with open(PLANILHA, "r+b"):
        pass
except PermissionError:
    print("❌ ERRO: Feche o arquivo 'automacao.xlsx' no Excel antes de rodar o script!")
    input("Pressione ENTER depois de fechar o arquivo...")

# =========================
# LENDO PLANILHA
# =========================
def normalizar(texto):
    return (
        unicodedata.normalize("NFKD", str(texto))
        .encode("ascii", errors="ignore")
        .decode("utf-8")
        .lower()
        .strip()
    )

df_raw = pd.read_excel(PLANILHA, sheet_name=SHEET, header=None, dtype=str)

header_row = None
for i, row in df_raw.iterrows():
    colunas = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip() != ""]
    if len(colunas) >= 5:
        header_row = i
        break

if header_row is None:
    print("❌ Não foi possível detectar o cabeçalho da planilha.")
    raise SystemExit

df = pd.read_excel(PLANILHA, sheet_name=SHEET, header=header_row, dtype=str)
df.columns = [normalizar(col) for col in df.columns]
df = df.dropna(how="all").reset_index(drop=True)

print("=" * 50)
print("Colunas encontradas na planilha:")
for i, col in enumerate(df.columns):
    print(f"  [{i}] '{col}'")
print(f"\nTotal de linhas: {len(df)}")
print("=" * 50)

if "chamado" not in df.columns:
    df["chamado"] = ""


def salvar_planilha():
    for tentativa in range(5):
        try:
            with pd.ExcelWriter(PLANILHA, mode="w", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=SHEET, index=False)
            print("  💾 Planilha salva.")
            return
        except PermissionError:
            if tentativa == 0:
                print("  ⚠️  Feche o 'automacao.xlsx' no Excel e pressione ENTER...")
                input()
            else:
                time.sleep(2)
    print("  ❌ Não foi possível salvar a planilha.")


# =========================
# DRIVER
# =========================
service = Service("C:/drivers/msedgedriver.exe")
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")

driver = webdriver.Edge(service=service, options=options)

# Timeouts centralizados — ajuste aqui se a rede estiver lenta
WAIT_LONGO  = WebDriverWait(driver, 120)  # login / submit
WAIT_MEDIO  = WebDriverWait(driver, 15)   # elementos normais do form
WAIT_CURTO  = WebDriverWait(driver, 5)    # elementos que devem já estar visíveis


# =========================
# FUNÇÕES AUXILIARES
# =========================

def aguardar_formulario():
    print("  ⏳ Aguardando formulário...")
    WAIT_LONGO.until(EC.presence_of_element_located((By.ID, "submit-btn")))
    WAIT_LONGO.until(EC.element_to_be_clickable((By.ID, "submit-btn")))
    # ✅ CORREÇÃO: pausa para os campos Select2 terminarem de inicializar
    # Sem isso, o campo "nome" pode ser clicado antes de estar pronto,
    # fazendo o sistema não retornar sugestões mesmo com nome correto.
    time.sleep(1.5)
    print("  ✅ Formulário pronto!")


def fechar_dropdown():
    """Fecha dropdown ativo somente se estiver aberto."""
    try:
        drops = driver.find_elements(By.CSS_SELECTOR, ".select2-drop-active")
        if not any(d.is_displayed() for d in drops):
            return
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        drops = driver.find_elements(By.CSS_SELECTOR, ".select2-drop-active")
        if any(d.is_displayed() for d in drops):
            driver.find_element(By.TAG_NAME, "body").click()
    except Exception:
        pass


def clicar(elemento):
    try:
        ActionChains(driver).move_to_element(elemento).click().perform()
    except Exception:
        driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", elemento)


def dropdown_aberto():
    try:
        drops = driver.find_elements(By.CSS_SELECTOR, ".select2-drop-active")
        return any(d.is_displayed() for d in drops)
    except Exception:
        return False


def abrir_select2(id_container_s2):
    """Tenta abrir o dropdown Select2 por 4 estratégias, sem sleeps fixos."""
    select_id = id_container_s2.replace("s2id_", "", 1)
    fechar_dropdown()

    # 1) Link padrão
    try:
        link = driver.find_element(By.CSS_SELECTOR, f"#{id_container_s2} a.select2-choice")
        clicar(link)
        if dropdown_aberto():
            return True
    except Exception:
        pass

    # 2) Container
    try:
        container = driver.find_element(By.ID, id_container_s2)
        clicar(container)
        if dropdown_aberto():
            return True
    except Exception:
        pass

    # 3) jQuery
    try:
        driver.execute_script(f'$("#{select_id}").select2("open");')
        if dropdown_aberto():
            return True
    except Exception:
        pass

    # 4) Focusser + espaço
    try:
        focusser = driver.find_element(By.CSS_SELECTOR, f"#{id_container_s2} input.select2-focusser")
        driver.execute_script("arguments[0].focus();", focusser)
        ActionChains(driver).send_keys_to_element(focusser, Keys.SPACE).perform()
        if dropdown_aberto():
            return True
    except Exception:
        pass

    return False


def obter_search_input():
    """Retorna o input de busca visível do dropdown aberto."""
    try:
        els = driver.find_elements(By.CSS_SELECTOR, "input[id^='s2id_autogen'][id$='_search']")
        visiveis = [e for e in els if e.is_displayed()]
        return visiveis[0] if visiveis else None
    except Exception:
        return None


def aguardar_opcoes_dropdown(timeout=5):
    """
    Aguarda opções aparecerem no dropdown sem sleep fixo.
    Retorna lista de elementos visíveis assim que encontrar.
    """
    seletores = [
        ".select2-results li.select2-result-selectable",
        ".select2-drop-active li.select2-result-selectable",
        ".select2-results li",
    ]
    fim = time.monotonic() + timeout
    while time.monotonic() < fim:
        for seletor in seletores:
            try:
                opcoes = driver.find_elements(By.CSS_SELECTOR, seletor)
                visiveis = [o for o in opcoes if o.is_displayed() and o.text.strip()]
                if visiveis:
                    return visiveis
            except Exception:
                pass
        time.sleep(0.08)  # poll mínimo para não travar a CPU
    return []


def aguardar_selecao_confirmada(id_container_s2, timeout=8):
    """
    Aguarda o Select2 exibir um valor selecionado (não vazio / não placeholder).
    Retorna True se confirmado dentro do timeout.
    """
    placeholders = {"selecione", "select", "", "selecione...", "select...", "carregando..."}
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.find_element(
                By.CSS_SELECTOR, f"#{id_container_s2} .select2-chosen"
            ).text.strip().lower() not in placeholders
        )
        valor = driver.find_element(
            By.CSS_SELECTOR, f"#{id_container_s2} .select2-chosen"
        ).text.strip()
        print(f"    ✅ Confirmado: '{valor}'")
        return True
    except Exception:
        print(f"    ⚠️  Seleção não confirmada em '{id_container_s2}', seguindo...")
        return False


def _digitar_no_search(search, texto):
    """Limpa e digita no campo de busca do Select2."""
    driver.execute_script("arguments[0].focus(); arguments[0].value = '';", search)
    search.send_keys(texto)


def _selecionar_por_celulas(visiveis, texto):
    """
    Match exato e depois parcial nas colunas das opções.
    Retorna True se selecionou.
    """
    texto_lower = texto.lower()

    for opcao in visiveis:
        partes = re.split(r'\s{2,}|\t', opcao.text.strip())
        colunas = [p.strip() for p in partes if p.strip()]
        if any(c.lower() == texto_lower for c in colunas):
            clicar(opcao)
            print(f"    ✅ Selecionado (exato): '{opcao.text.strip()}'")
            return True

    for opcao in visiveis:
        partes = re.split(r'\s{2,}|\t', opcao.text.strip())
        colunas = [p.strip() for p in partes if p.strip()]
        if any(texto_lower in c.lower() for c in colunas):
            clicar(opcao)
            print(f"    ✅ Selecionado (parcial): '{opcao.text.strip()}'")
            return True

    return False


# =========================
# FUNÇÕES DE PREENCHIMENTO
# =========================

def selecionar_select2(id_container_s2, texto_opcao):
    texto_opcao = str(texto_opcao).strip()

    try:
        WAIT_CURTO.until(EC.visibility_of_element_located((By.ID, id_container_s2)))
    except Exception:
        pass

    if not abrir_select2(id_container_s2):
        print(f"    ❌ Não conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search:
        _digitar_no_search(search, texto_opcao)

    visiveis = aguardar_opcoes_dropdown(timeout=4)
    selecionou = False

    if visiveis:
        for opcao in visiveis:
            if opcao.text.strip().lower() == texto_opcao.lower():
                clicar(opcao)
                print(f"    ✅ Selecionado: '{opcao.text.strip()}'")
                selecionou = True
                break
        if not selecionou:
            clicar(visiveis[0])
            print(f"    ⚠️  '{texto_opcao}' não encontrado — selecionou '{visiveis[0].text.strip()}'")
            selecionou = True
    else:
        print(f"    ❌ Nenhuma opção visível para '{texto_opcao}'")

    fechar_dropdown()
    if selecionou:
        aguardar_selecao_confirmada(id_container_s2)


def preencher_reference(id_container_s2, texto):
    texto = str(texto).strip()

    if not abrir_select2(id_container_s2):
        print(f"    ❌ Não conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search is None:
        print(f"    ❌ Campo de busca não apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return

    _digitar_no_search(search, texto)
    visiveis = aguardar_opcoes_dropdown(timeout=5)

    selecionou = False
    if visiveis:
        selecionou = _selecionar_por_celulas(visiveis, texto)
        if not selecionou:
            clicar(visiveis[0])
            print(f"    ⚠️  '{texto}' não encontrado — selecionou primeiro: '{visiveis[0].text.strip()}'")
            selecionou = True
    else:
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        print(f"    ⚠️  Sem sugestão para '{texto}'")

    fechar_dropdown()
    if selecionou:
        aguardar_selecao_confirmada(id_container_s2)


def preencher_reference_local(id_container_s2, texto):
    """
    Digita o texto no campo de busca e seleciona SEMPRE a primeira opção
    que aparecer, sem tentar correspondência exata.
    """
    texto = str(texto).strip()

    if not abrir_select2(id_container_s2):
        print(f"    ❌ Não conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search is None:
        print(f"    ❌ Campo de busca não apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return

    _digitar_no_search(search, texto)
    time.sleep(1.5)

    visiveis = aguardar_opcoes_dropdown(timeout=8)

    if not visiveis:
        print(f"    ⚠️  Nenhuma sugestão para '{texto}' no campo local")
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        fechar_dropdown()
        return

    primeira = visiveis[0]
    print(f"    → Primeira opção disponível: '{primeira.text.strip()}'")
    clicar(primeira)
    print(f"    ✅ Local selecionado: '{primeira.text.strip()}'")

    fechar_dropdown()
    aguardar_selecao_confirmada(id_container_s2)


def preencher_reference_nome(id_container_s2, texto):
    """
    Tenta selecionar a pessoa pelo nome ou e-mail.
    Retorna True se encontrou e selecionou, False se não encontrou ninguém.
    """
    texto = str(texto).strip()

    if not abrir_select2(id_container_s2):
        print(f"    ❌ Não conseguiu abrir '{id_container_s2}'")
        return False

    search = obter_search_input()
    if search is None:
        print(f"    ❌ Campo de busca não apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return False

    _digitar_no_search(search, texto)

    
    visiveis = aguardar_opcoes_dropdown(timeout=8)

    if not visiveis:
     
        print(f"    ⚠️  Sem sugestões na 1ª tentativa — tentando novamente...")
        _digitar_no_search(search, texto)
        time.sleep(1.0)
        visiveis = aguardar_opcoes_dropdown(timeout=8)

    if not visiveis:
        print(f"    ❌ Nome '{texto}' não encontrado no sistema (sem sugestões)")
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        fechar_dropdown()
        return False

    selecionou = False
    texto_lower = texto.lower()

    # Match pelo nome (coluna 1)
    for li in visiveis:
        try:
            celulas = li.find_elements(By.CSS_SELECTOR, "div.select2-result-cell")
            col1 = celulas[0].text.strip() if celulas else ""
            if col1.lower() == texto_lower:
                clicar(li)
                print(f"    ✅ Selecionado pelo nome: '{col1}'")
                selecionou = True
                break
        except Exception:
            continue

    # Match pelo e-mail (coluna 2)
    if not selecionou:
        for li in visiveis:
            try:
                celulas = li.find_elements(By.CSS_SELECTOR, "div.select2-result-cell")
                col2 = celulas[1].text.strip() if len(celulas) >= 2 else ""
                prefixo = col2.split("@")[0] if "@" in col2 else col2
                if prefixo.lower() == texto_lower:
                    clicar(li)
                    print(f"    ✅ Selecionado pelo e-mail: '{col2}'")
                    selecionou = True
                    break
            except Exception:
                continue

    if not selecionou:
        resumo = []
        for li in visiveis[:5]:
            try:
                celulas = li.find_elements(By.CSS_SELECTOR, "div.select2-result-cell")
                c1 = celulas[0].text.strip() if celulas else ""
                c2 = celulas[1].text.strip() if len(celulas) >= 2 else ""
                resumo.append(f"[{c1} | {c2}]")
            except Exception:
                resumo.append(li.text.strip())
        print(f"    ❌ Sem correspondência exata para '{texto}'. Opções: {resumo}")
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        fechar_dropdown()
        return False

    fechar_dropdown()
    aguardar_selecao_confirmada(id_container_s2)
    return True


def preencher_reference_setor(id_container_s2, texto):
    """Campo setor: tenta match exato primeiro. Se não achar, digita 'adm' e seleciona a 1ª opção."""
    texto = str(texto).strip()

    if not abrir_select2(id_container_s2):
        print(f"    ❌ Não conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search is None:
        print(f"    ❌ Campo de busca não apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return

    _digitar_no_search(search, texto)
    visiveis = aguardar_opcoes_dropdown(timeout=4)
    selecionou = _selecionar_por_celulas(visiveis, texto) if visiveis else False

    if not selecionou:
        print(f"    ⚠️  '{texto}' não encontrado — fallback: digitando 'adm' e selecionando 1ª opção...")
        _digitar_no_search(search, "adm")
        time.sleep(1.0)
        visiveis = aguardar_opcoes_dropdown(timeout=6)

        if visiveis:
            primeira = visiveis[0]
            print(f"    → Primeira opção disponível: '{primeira.text.strip()}'")
            clicar(primeira)
            print(f"    ✅ Setor selecionado (fallback): '{primeira.text.strip()}'")
            selecionou = True
        else:
            print(f"    ❌ Nenhuma opção encontrada nem com 'adm'.")
            try:
                search.send_keys(Keys.ESCAPE)
            except Exception:
                pass

    fechar_dropdown()
    if selecionou:
        aguardar_selecao_confirmada(id_container_s2)


def preencher_input(id_campo, valor):
    fechar_dropdown()
    el = WAIT_MEDIO.until(EC.element_to_be_clickable((By.ID, id_campo)))
    el.clear()
    el.send_keys(str(valor).strip())


def capturar_protocolo():
    """
    Captura o número RITM/REQ após envio do formulário.
    Tenta 4 estratégias em ordem crescente de complexidade.
    """

    try:
        WebDriverWait(driver, 10).until(
            lambda d: re.search(r"(RITM|REQ|INC|CHG)\d+", d.title, re.IGNORECASE)
        )
        match = re.search(r"(RITM|REQ|INC|CHG)\d+", driver.title, re.IGNORECASE)
        if match:
            print(f"  📋 Protocolo (título): {match.group(0)}")
            return match.group(0)
    except Exception:
        pass

   
    seletores_ritm = [
        "td.wrapper-md a[title^='RITM']",
        "td.wrapper-md a[title^='REQ']",
        "td.wrapper-md a[href*='sc_req_item']",
        "a[title^='RITM']",
        "a[title^='REQ']",
    ]
    for seletor in seletores_ritm:
        try:
            WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.CSS_SELECTOR, seletor)))
            el = driver.find_element(By.CSS_SELECTOR, seletor)
            numero = el.get_attribute("title") or el.text.strip()
            if re.match(r"(RITM|REQ|INC|CHG)\d+", numero, re.IGNORECASE):
                print(f"  📋 Protocolo: {numero}")
                return numero
        except Exception:
            continue

   
    seletores_texto = [
        "h1.page-header", "div.panel-heading h1", "span.outputmsg_text",
        "div#outputmsg", "div.request-number", "h2",
        ".form-group .form-control-static",
    ]
    for seletor in seletores_texto:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, seletor):
                match = re.search(r"(RITM|REQ|INC|CHG|TASK)\d+", el.text.strip(), re.IGNORECASE)
                if match:
                    return match.group(0)
        except Exception:
            continue

    
    match = re.search(r"(RITM|REQ|INC)\d+", driver.current_url, re.IGNORECASE)
    if match:
        return match.group(0)

    print("  ⚠️  ATENÇÃO: chamado enviado mas número não capturado automaticamente.")
    print("              Verifique manualmente no ServiceNow e corrija a planilha.")
    return "N/A"


# =========================
# MAPEAMENTO DOS CAMPOS
# =========================

SELECT2 = {
    "usuario":             "s2id_sp_formfield_u_solicitacao_propria",
    "localizado":          "s2id_sp_formfield_u_user_located",
    "gerente: s/n":        "s2id_sp_formfield_u_confirmar_gestor",
    "servico ti":          "s2id_sp_formfield_qual_servico_tecnologia",
    "tipo de solicitacao": "s2id_sp_formfield_qual_servico_microinformatica",
    "tipo de servico":     "s2id_sp_formfield_tipo_de_servi_o_microinformatica_Programas",
    "aplicativo":          "s2id_sp_formfield_selecione_o_programa_aplicativo",
    "utilizado":           "s2id_sp_formfield_o_software_aplicativo_ser_utilizado_pelo_pr_prio_solicitante",
}

REFERENCE_NOME = "s2id_sp_formfield_requested_for"

REFERENCE = {}

INPUTS = {
    "telefone":    "sp_formfield_u_informe_telefone_para_contato",
    "setor inst.": "sp_formfield_setor_de_instala_o_do_software_aplicativo",
    "descricao":   "sp_formfield_description",
}

ORDEM = [
    "usuario", "nome", "localizado", "telefone", "gerente: s/n",
    "local", "setor", "servico ti", "tipo de solicitacao",
    "tipo de servico", "aplicativo", "utilizado", "setor inst.", "descricao"
]

LOCAL_S2_ID = "s2id_sp_formfield_u_unidade_hosp"


def valor_valido(val):
    return pd.notna(val) and str(val).strip().lower() not in ("", "nan", "none")


# =========================
# LOOP PRINCIPAL
# =========================

linhas_sem_nome = []
linhas_sem_numero = []  

primeira_pendente = None
for index, linha in df.iterrows():
    chamado_vazio = not valor_valido(linha.get("chamado", ""))
    nome_vazio    = not valor_valido(linha.get("nome", ""))

    if chamado_vazio and not nome_vazio:
        primeira_pendente = index
        break

if primeira_pendente is None:
    print("✅ Todas as linhas pendentes já possuem chamado ou estão sem nome. Nada a fazer.")
    driver.quit()
    raise SystemExit

print(f"\n▶ Iniciando a partir da linha {primeira_pendente + 1}\n")

for index, linha in df.iterrows():
    chamado_atual = str(linha.get("chamado", "")).strip()
    chamado_vazio = not valor_valido(linha.get("chamado", ""))
    nome_vazio    = not valor_valido(linha.get("nome", ""))

    
    chamado_pendente = chamado_vazio or chamado_atual == "N/A"

    if not chamado_pendente:
        print(f"⏭ Linha {index + 1} já tem chamado ({chamado_atual}), pulando.")
        continue

    if nome_vazio:
        print(f"⏭ Linha {index + 1} sem nome preenchido, pulando.")
        continue

    nome_da_linha = str(linha.get("nome", "")).strip()
    reprocessando = chamado_atual == "N/A"
    status_label = "🔁 REPROCESSANDO (era N/A)" if reprocessando else "🔵 Nova solicitação"
    print(f"\n{status_label} {index + 1}/{len(df)} — Nome: {nome_da_linha}")

    try:
        driver.get(URL)
        aguardar_formulario()
    except Exception as e:
        print(f"  ❌ Erro ao carregar formulário: {e}")
        continue

    nome_encontrado = True

    for col in ORDEM:
        if col not in df.columns:
            continue
        valor = linha.get(col, "")
        if not valor_valido(valor):
            continue

        print(f"  → '{col}' = '{valor}'")
        try:
            if col == "nome":
                nome_encontrado = preencher_reference_nome(REFERENCE_NOME, valor)

                if not nome_encontrado:
                    print(f"  ⚠️  Nome '{valor}' não encontrado. Pulando linha {index + 1}...")
                    linhas_sem_nome.append({
                        "linha":  index + 1,
                        "nome":   valor,
                        "motivo": "Nome não encontrado no sistema ServiceNow"
                    })
                    try:
                        driver.get("about:blank")
                    except Exception:
                        pass
                    break

            elif col == "setor":
                preencher_reference_setor("s2id_sp_formfield_u_setor", valor)

            elif col == "local":
                preencher_reference_local(LOCAL_S2_ID, valor)

            elif col in SELECT2:
                selecionar_select2(SELECT2[col], valor)
                if col == "tipo de servico":
                    print("  ⏳ Aguardando campo aplicativo...")
                    try:
                        WAIT_CURTO.until(EC.visibility_of_element_located(
                            (By.ID, "s2id_sp_formfield_selecione_o_programa_aplicativo")
                        ))
                        print("  ✅ Campo aplicativo pronto!")
                    except Exception:
                        pass

            elif col in REFERENCE:
                preencher_reference(REFERENCE[col], valor)

            elif col in INPUTS:
                preencher_input(INPUTS[col], valor)

        except Exception as e:
            print(f"    ❌ Erro no campo '{col}': {e}")
            fechar_dropdown()

    if not nome_encontrado:
        continue

    print("  → Enviando formulário...")
    try:
        submit = WAIT_LONGO.until(EC.element_to_be_clickable((By.ID, "submit-btn")))
        clicar(submit)
    except Exception as e:
        print(f"  ❌ Erro ao enviar: {e}")
        continue

    protocolo = capturar_protocolo()
    df.at[index, "chamado"] = protocolo
    print(f"  ✅ Chamado: {protocolo}")

    
    if protocolo == "N/A":
        linhas_sem_numero.append({"linha": index + 1, "nome": nome_da_linha})

    salvar_planilha()

    if index < len(df) - 1:
        print("  🔄 Próximo chamado...")
        driver.back()
        try:
            aguardar_formulario()
        except Exception:
            driver.get(URL)
            aguardar_formulario()

driver.quit()
print("\n🎉 Concluído! Chamados salvos em 'automacao.xlsx'.")

# =========================
# RELATÓRIO FINAL
# =========================
print("\n" + "=" * 60)

if linhas_sem_nome:
    print(f"⚠️  RELATÓRIO — {len(linhas_sem_nome)} chamado(s) NÃO aberto(s) por nome não encontrado:\n")
    print(f"  {'Linha':<8} {'Nome'}")
    print(f"  {'-'*8} {'-'*40}")
    for item in linhas_sem_nome:
        print(f"  {item['linha']:<8} {item['nome']}")
    print(f"\n  Motivo: {linhas_sem_nome[0]['motivo']}")
    print("\n  💡 Verifique se os nomes estão cadastrados no ServiceNow")
    print("     e corrija a planilha para reprocessar estas linhas.")
else:
    print("✅ Todos os chamados foram abertos com sucesso! Nenhum nome faltou.")

if linhas_sem_numero:
    print(f"\n⚠️  RELATÓRIO — {len(linhas_sem_numero)} chamado(s) aberto(s) mas SEM número capturado (N/A):\n")
    print(f"  {'Linha':<8} {'Nome'}")
    print(f"  {'-'*8} {'-'*40}")
    for item in linhas_sem_numero:
        print(f"  {item['linha']:<8} {item['nome']}")
    print("\n  💡 O formulário FOI enviado, mas o número não apareceu a tempo.")
    print("     Acesse o ServiceNow, filtre por data/hora e corrija a planilha.")
    print("     Na próxima execução, linhas com N/A serão reprocessadas automaticamente.")

print("=" * 60)

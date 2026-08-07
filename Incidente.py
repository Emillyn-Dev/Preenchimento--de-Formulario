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
# CONFIGURACAO
# =========================
# Ajuste os dois caminhos abaixo antes de rodar na empresa
os.chdir(r"C:\Levar\Formulario") #caminho da pasta do código

URL      = "x"  #url do site que vai rodar o código
PLANILHA = "RITM - INC.xlsx" #nome da planilha excel
SHEET    = "INC" #planilha que quero consultar pra abertura do chamado

# =========================
# VERIFICA SE PLANILHA ESTA ABERTA
# =========================
try:
    with open(PLANILHA, "r+b"):
        pass
except PermissionError:
    print("ERRO: Feche o arquivo no Excel antes de rodar o script!")
    input("Pressione ENTER depois de fechar o arquivo...")

# =========================
# LEITURA DA PLANILHA
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
    if len(colunas) >= 3:
        header_row = i
        break

if header_row is None:
    print("Nao foi possivel detectar o cabecalho da planilha.")
    raise SystemExit

df = pd.read_excel(PLANILHA, sheet_name=SHEET, header=header_row, dtype=str)
df.columns = [normalizar(col) for col in df.columns]
df = df.dropna(how="all").reset_index(drop=True)

if "chamado" not in df.columns:
    df["chamado"] = ""

print("=" * 50)
print("Colunas encontradas na planilha INC:")
for i, col in enumerate(df.columns):
    print(f"  [{i}] '{col}'")
print(f"\nTotal de linhas: {len(df)}")
print("=" * 50)


def salvar_planilha():
    for tentativa in range(5):
        try:
            with pd.ExcelWriter(PLANILHA, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=SHEET, index=False)
            print("  Planilha salva.")
            return
        except PermissionError:
            if tentativa == 0:
                print("  Feche o arquivo no Excel e pressione ENTER...")
                input()
            else:
                time.sleep(2)
    print("  Nao foi possivel salvar a planilha.")


# =========================
# DRIVER
# =========================
# Ajuste o caminho do msedgedriver.exe abaixo
service = Service("x") #caminho 
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")

driver = webdriver.Edge(service=service, options=options)

WAIT_LONGO = WebDriverWait(driver, 120)
WAIT_MEDIO = WebDriverWait(driver, 15)
WAIT_CURTO = WebDriverWait(driver, 5)


# =========================
# FUNCOES AUXILIARES
# =========================

def aguardar_formulario():
    print("  Aguardando formulario...")
    WAIT_LONGO.until(EC.presence_of_element_located((By.ID, "submit-btn")))
    WAIT_LONGO.until(EC.element_to_be_clickable((By.ID, "submit-btn")))
    time.sleep(1.5)
    print("  Formulario pronto!")


def fechar_dropdown():
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
    select_id = id_container_s2.replace("s2id_", "", 1)
    fechar_dropdown()

    try:
        link = driver.find_element(By.CSS_SELECTOR, f"#{id_container_s2} a.select2-choice")
        clicar(link)
        if dropdown_aberto():
            return True
    except Exception:
        pass

    try:
        container = driver.find_element(By.ID, id_container_s2)
        clicar(container)
        if dropdown_aberto():
            return True
    except Exception:
        pass

    try:
        driver.execute_script(f'$("#{select_id}").select2("open");')
        if dropdown_aberto():
            return True
    except Exception:
        pass

    try:
        focusser = driver.find_element(By.CSS_SELECTOR, f"#{id_container_s2} input.select2-focusser")
        driver.execute_script("arguments[0].focus();", focusser)
        ActionChains(driver).send_keys_to_element(focusser, Keys.SPACE).perform()
        if dropdown_aberto():
            return True
    except Exception:
        pass

    return False


def abrir_select2_nome():
    """
    O campo de nome tem ID Angular dinamico (s2id_sp_formfield_{{::field.name}}).
    Abrimos pelo aria-label ou eliminando os campos ja conhecidos.
    """
    fechar_dropdown()

   
    IGNORAR = {
        "s2id_sp_formfield_location",
        "s2id_sp_formfield_u_sector",
        "s2id_sp_formfield_category",
        "s2id_sp_formfield_service_offering",
        "s2id_sp_formfield_urgency",
    }

    try:
        # Tenta pelo aria-label do link
        links = driver.find_elements(By.CSS_SELECTOR, "a.select2-choice")
        for link in links:
            aria = link.get_attribute("aria-label") or ""
            if "pessoa" in aria.lower() or "atendida" in aria.lower() or "caller" in aria.lower():
                clicar(link)
                if dropdown_aberto():
                    return True

        
        containers = driver.find_elements(By.CSS_SELECTOR, "div.select2-container")
        for c in containers:
            cid = c.get_attribute("id") or ""
            if cid in IGNORAR:
                continue
            try:
                link = c.find_element(By.CSS_SELECTOR, "a.select2-choice")
                clicar(link)
                if dropdown_aberto():
                    return True
                fechar_dropdown()
            except Exception:
                continue
    except Exception:
        pass

    return False


def obter_search_input():
    try:
        els = driver.find_elements(By.CSS_SELECTOR, "input[id^='s2id_autogen'][id$='_search']")
        visiveis = [e for e in els if e.is_displayed()]
        return visiveis[0] if visiveis else None
    except Exception:
        return None


def aguardar_opcoes_dropdown(timeout=5):
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
        time.sleep(0.08)
    return []


def aguardar_opcoes_estabilizadas(timeout=10, janela_estavel=0.8, intervalo=0.25):
    fim = time.monotonic() + timeout
    visiveis = aguardar_opcoes_dropdown(timeout=timeout)
    if not visiveis:
        return []

    def assinatura(opcoes):
        return tuple(o.text.strip() for o in opcoes)

    anterior = assinatura(visiveis)
    desde_quando_estavel = time.monotonic()

    while time.monotonic() < fim:
        time.sleep(intervalo)
        try:
            opcoes = driver.find_elements(By.CSS_SELECTOR, ".select2-results li.select2-result-selectable")
            visiveis = [o for o in opcoes if o.is_displayed() and o.text.strip()]
        except Exception:
            visiveis = []

        if not visiveis:
            anterior = ()
            desde_quando_estavel = time.monotonic()
            continue

        atual = assinatura(visiveis)
        if atual != anterior:
            anterior = atual
            desde_quando_estavel = time.monotonic()
        else:
            if time.monotonic() - desde_quando_estavel >= janela_estavel:
                return visiveis

    return visiveis


def aguardar_selecao_confirmada(id_container_s2, timeout=8):
    placeholders = {"selecione", "select", "", "selecione...", "select...", "carregando..."}
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.find_element(
                By.CSS_SELECTOR, f"#{id_container_s2} .select2-chosen"
            ).text.strip().lower() not in placeholders
        )
        valor = driver.find_element(By.CSS_SELECTOR, f"#{id_container_s2} .select2-chosen").text.strip()
        print(f"    Confirmado: '{valor}'")
        return True
    except Exception:
        print(f"    Selecao nao confirmada em '{id_container_s2}', seguindo...")
        return False


def _digitar_no_search(search, texto):
    driver.execute_script("arguments[0].focus(); arguments[0].value = '';", search)
    search.send_keys(texto)


def _selecionar_por_celulas(visiveis, texto):
    texto_lower = texto.lower()
    for opcao in visiveis:
        partes = re.split(r'\s{2,}|\t', opcao.text.strip())
        colunas = [p.strip() for p in partes if p.strip()]
        if any(c.lower() == texto_lower for c in colunas):
            clicar(opcao)
            print(f"    Selecionado (exato): '{opcao.text.strip()}'")
            return True
    for opcao in visiveis:
        partes = re.split(r'\s{2,}|\t', opcao.text.strip())
        colunas = [p.strip() for p in partes if p.strip()]
        if any(texto_lower in c.lower() for c in colunas):
            clicar(opcao)
            print(f"    Selecionado (parcial): '{opcao.text.strip()}'")
            return True
    return False


# =========================
# FUNCOES DE PREENCHIMENTO
# =========================

def preencher_nome_inc(texto):
    """
    Preenche o campo de nome do formulario INC.
    O ID do container e dinamico (Angular), entao usa abrir_select2_nome().
    Retorna True se selecionou, False se nao encontrou.
    """
    texto = str(texto).strip()

    if not abrir_select2_nome():
        print(f"    Nao conseguiu abrir o campo de nome")
        return False

    search = obter_search_input()
    if search is None:
        print(f"    Campo de busca nao apareceu para nome")
        fechar_dropdown()
        return False

    _digitar_no_search(search, texto)

    print("    Aguardando o sistema comecar a buscar...")
    time.sleep(2.0)

    print("    Aguardando lista estabilizar...")
    visiveis = aguardar_opcoes_estabilizadas(timeout=12, janela_estavel=0.8)

    if not visiveis:
        print(f"    Sem sugestoes na 1a tentativa -- tentando novamente...")
        _digitar_no_search(search, texto)
        time.sleep(2.0)
        visiveis = aguardar_opcoes_estabilizadas(timeout=12, janela_estavel=0.8)

    if not visiveis:
        print(f"    Nome '{texto}' nao encontrado (sem sugestoes)")
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
            col1 = celulas[0].text.strip() if celulas else li.text.strip()
            if col1.lower() == texto_lower:
                clicar(li)
                print(f"    Selecionado pelo nome: '{col1}'")
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
                    print(f"    Selecionado pelo e-mail: '{col2}'")
                    selecionou = True
                    break
            except Exception:
                continue

    if not selecionou:
        resumo = []
        for li in visiveis[:5]:
            try:
                celulas = li.find_elements(By.CSS_SELECTOR, "div.select2-result-cell")
                c1 = celulas[0].text.strip() if celulas else li.text.strip()
                c2 = celulas[1].text.strip() if len(celulas) >= 2 else ""
                resumo.append(f"[{c1} | {c2}]")
            except Exception:
                resumo.append(li.text.strip())
        print(f"    Sem correspondencia exata para '{texto}'. Opcoes: {resumo}")
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        fechar_dropdown()
        return False

    fechar_dropdown()

    # Confirmacao: verifica se algum select2-chosen exibe o nome
    try:
        WebDriverWait(driver, 8).until(
            lambda d: any(
                texto_lower in el.text.strip().lower()
                for el in d.find_elements(By.CSS_SELECTOR, ".select2-chosen")
                if el.is_displayed() and el.text.strip()
            )
        )
        print(f"    Nome confirmado no campo.")
    except Exception:
        print(f"    Confirmacao de nome nao detectada, seguindo...")
    return True


def preencher_reference_local(id_container_s2, texto):
    """Digita o texto e seleciona sempre a primeira opcao disponivel."""
    texto = str(texto).strip()

    if not abrir_select2(id_container_s2):
        print(f"    Nao conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search is None:
        print(f"    Campo de busca nao apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return

    _digitar_no_search(search, texto)
    time.sleep(1.5)

    visiveis = aguardar_opcoes_dropdown(timeout=8)

    if not visiveis:
        print(f"    Nenhuma sugestao para '{texto}' no campo local")
        try:
            search.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        fechar_dropdown()
        return

    primeira = visiveis[0]
    print(f"    Primeira opcao disponivel: '{primeira.text.strip()}'")
    clicar(primeira)
    print(f"    Local selecionado: '{primeira.text.strip()}'")

    fechar_dropdown()
    aguardar_selecao_confirmada(id_container_s2)


def preencher_reference_setor(id_container_s2, texto):
    """
    Digita o texto no campo de busca do setor, aguarda a lista filtrar
    e seleciona o item com match exato.
    O campo filtra as opcoes conforme o usuario digita (como visto na imagem).
    """
    texto = str(texto).strip()
    texto_lower = texto.lower()

    if not abrir_select2(id_container_s2):
        print(f"    Nao conseguiu abrir '{id_container_s2}'")
        return

    search = obter_search_input()
    if search is None:
        print(f"    Campo de busca nao apareceu para '{id_container_s2}'")
        fechar_dropdown()
        return

    # Cola o texto e aguarda a lista filtrar
    _digitar_no_search(search, texto)
    print(f"    Aguardando lista filtrar por '{texto}'...")
    time.sleep(1.5)

    visiveis = aguardar_opcoes_estabilizadas(timeout=8, janela_estavel=0.6)

    if not visiveis:
        print(f"    Nenhuma opcao apareceu apos filtrar por '{texto}' -- tentando com 'adm'...")
        _digitar_no_search(search, "adm")
        time.sleep(1.5)
        visiveis = aguardar_opcoes_estabilizadas(timeout=8, janela_estavel=0.6)

        if not visiveis:
            print(f"    Nenhuma opcao apareceu mesmo buscando por 'adm'")
            try:
                search.send_keys(Keys.ESCAPE)
            except Exception:
                pass
            fechar_dropdown()
            return

        clicar(visiveis[0])
        print(f"    Setor selecionado (fallback 'adm'): '{visiveis[0].text.strip()}'")
        fechar_dropdown()
        aguardar_selecao_confirmada(id_container_s2)
        return

    selecionou = False

    # Match exato
    for opcao in visiveis:
        if opcao.text.strip().lower() == texto_lower:
            clicar(opcao)
            print(f"    Setor selecionado (exato): '{opcao.text.strip()}'")
            selecionou = True
            break

    # Match parcial
    if not selecionou:
        for opcao in visiveis:
            if texto_lower in opcao.text.strip().lower():
                clicar(opcao)
                print(f"    Setor selecionado (parcial): '{opcao.text.strip()}'")
                selecionou = True
                break

    # Fallback: primeira opcao da lista filtrada
    if not selecionou:
        print(f"    '{texto}' nao encontrado -- selecionando primeiro da lista filtrada...")
        clicar(visiveis[0])
        print(f"    Setor selecionado (fallback): '{visiveis[0].text.strip()}'")
        selecionou = True

    fechar_dropdown()
    if selecionou:
        aguardar_selecao_confirmada(id_container_s2)


def preencher_input(id_campo, valor):
    fechar_dropdown()
    el = WAIT_MEDIO.until(EC.element_to_be_clickable((By.ID, id_campo)))
    el.clear()
    el.send_keys(str(valor).strip())


def preencher_textarea(id_campo, valor):
    fechar_dropdown()
    el = WAIT_MEDIO.until(EC.element_to_be_clickable((By.ID, id_campo)))
    el.clear()
    el.send_keys(str(valor).strip())


def capturar_protocolo():
    """Captura o numero INC apos envio. 4 estrategias em ordem crescente."""

    
    try:
        WebDriverWait(driver, 25).until(
            lambda d: any(
                re.search(r"INC\d+", el.text.strip(), re.IGNORECASE)
                for el in d.find_elements(By.CSS_SELECTOR, "span.pre-wrap.ng-binding")
                if el.is_displayed()
            )
        )
        for el in driver.find_elements(By.CSS_SELECTOR, "span.pre-wrap.ng-binding"):
            match = re.search(r"INC\d+", el.text.strip(), re.IGNORECASE)
            if match and el.is_displayed():
                print(f"  Protocolo (span): {match.group(0)}")
                return match.group(0)
    except Exception:
        pass

  
    try:
        WebDriverWait(driver, 10).until(
            lambda d: re.search(r"(INC|REQ|RITM|CHG)\d+", d.title, re.IGNORECASE)
        )
        match = re.search(r"(INC|REQ|RITM|CHG)\d+", driver.title, re.IGNORECASE)
        if match:
            print(f"  Protocolo (titulo): {match.group(0)}")
            return match.group(0)
    except Exception:
        pass

    
    seletores_texto = [
        "span.ng-binding", "h1.page-header", "div.panel-heading h1",
        "span.outputmsg_text", "div#outputmsg", "div.request-number",
        "h2", ".form-group .form-control-static",
    ]
    for seletor in seletores_texto:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, seletor):
                match = re.search(r"(INC|REQ|RITM|CHG|TASK)\d+", el.text.strip(), re.IGNORECASE)
                if match and el.is_displayed():
                    print(f"  Protocolo (texto): {match.group(0)}")
                    return match.group(0)
        except Exception:
            continue

    
    match = re.search(r"(INC|REQ|RITM)\d+", driver.current_url, re.IGNORECASE)
    if match:
        print(f"  Protocolo (URL): {match.group(0)}")
        return match.group(0)

    print("  ATENCAO: chamado enviado mas numero nao capturado automaticamente.")
    print("           Verifique manualmente no ServiceNow e corrija a planilha.")
    return "N/A"


# =========================
# MAPEAMENTO DOS CAMPOS -- INC (Relatar um Erro)
# IDs confirmados via console do F12
# =========================

# Campo "Local / Unidade" -- confirmado
LOCAL_S2_ID = "s2id_sp_formfield_location"

# Campo "Setor" -- confirmado
SETOR_S2_ID = "s2id_sp_formfield_u_sector"

# Campo "Telefone" -- confirmar via F12 se nao funcionar
INPUT_TELEFONE_ID = "sp_formfield_u_contact_phone"

# Campo "Descricao / Descreva seu erro" -- confirmar via F12 se nao funcionar
TEXTAREA_DESC_ID = "sp_formfield_comments"

# Ordem de preenchimento dos campos
ORDEM = ["nome", "telefone", "local", "setor", "descricao"]


def valor_valido(val):
    return pd.notna(val) and str(val).strip().lower() not in ("", "nan", "none")


# =========================
# LOOP PRINCIPAL
# =========================

linhas_sem_nome   = []
linhas_sem_numero = []

# Encontra a primeira linha pendente
primeira_pendente = None
for index, linha in df.iterrows():
    chamado_vazio = not valor_valido(linha.get("chamado", ""))
    nome_vazio    = not valor_valido(linha.get("nome", ""))
    if chamado_vazio and not nome_vazio:
        primeira_pendente = index
        break

if primeira_pendente is None:
    print("Todas as linhas pendentes ja possuem chamado ou estao sem nome. Nada a fazer.")
    driver.quit()
    raise SystemExit

print(f"\nIniciando a partir da linha {primeira_pendente + 1}\n")

for index, linha in df.iterrows():
    chamado_atual    = str(linha.get("chamado", "")).strip()
    chamado_vazio    = not valor_valido(linha.get("chamado", ""))
    nome_vazio       = not valor_valido(linha.get("nome", ""))
    chamado_pendente = chamado_vazio or chamado_atual == "N/A"

    if not chamado_pendente:
        print(f"Linha {index + 1} ja tem chamado ({chamado_atual}), pulando.")
        continue

    if nome_vazio:
        print(f"Linha {index + 1} sem nome preenchido, pulando.")
        continue

    nome_da_linha = str(linha.get("nome", "")).strip()
    reprocessando = chamado_atual == "N/A"
    status_label  = "REPROCESSANDO (era N/A)" if reprocessando else "Nova solicitacao"
    print(f"\n{status_label} -- Linha {index + 1}/{len(df)} -- Nome: {nome_da_linha}")

    try:
        driver.get(URL)
        aguardar_formulario()
    except Exception as e:
        print(f"  Erro ao carregar formulario: {e}")
        continue

    nome_encontrado = True

    for col in ORDEM:
        if col not in df.columns:
            continue
        valor = linha.get(col, "")
        if not valor_valido(valor):
            continue

        print(f"  Campo '{col}' = '{valor}'")
        try:
            if col == "nome":
                nome_encontrado = preencher_nome_inc(valor)
                if not nome_encontrado:
                    print(f"  Nome '{valor}' nao encontrado. Pulando linha {index + 1}...")
                    linhas_sem_nome.append({
                        "linha":  index + 1,
                        "nome":   valor,
                        "motivo": "Nome nao encontrado no sistema ServiceNow"
                    })
                    try:
                        driver.get("about:blank")
                    except Exception:
                        pass
                    break

            elif col == "local":
                preencher_reference_local(LOCAL_S2_ID, valor)

            elif col == "setor":
                preencher_reference_setor(SETOR_S2_ID, valor)

            elif col == "telefone":
                preencher_input(INPUT_TELEFONE_ID, valor)

            elif col == "descricao":
                preencher_textarea(TEXTAREA_DESC_ID, valor)

        except Exception as e:
            print(f"    Erro no campo '{col}': {e}")
            fechar_dropdown()

    if not nome_encontrado:
        continue

    print("  Enviando formulario...")
    try:
        submit = WAIT_LONGO.until(EC.element_to_be_clickable((By.ID, "submit-btn")))
        clicar(submit)
    except Exception as e:
        print(f"  Erro ao enviar: {e}")
        continue

    protocolo = capturar_protocolo()
    df.at[index, "chamado"] = protocolo
    print(f"  Chamado: {protocolo}")

    if protocolo == "N/A":
        linhas_sem_numero.append({"linha": index + 1, "nome": nome_da_linha})

    salvar_planilha()

    if index < len(df) - 1:
        print("  Proximo chamado...")
        driver.back()
        try:
            aguardar_formulario()
        except Exception:
            driver.get(URL)
            aguardar_formulario()

driver.quit()
print("\nConcluido! Chamados INC salvos na aba 'INC' da planilha.")

# =========================
# RELATORIO FINAL
# =========================
print("\n" + "=" * 60)

if linhas_sem_nome:
    print(f"RELATORIO -- {len(linhas_sem_nome)} chamado(s) NAO aberto(s) por nome nao encontrado:\n")
    print(f"  {'Linha':<8} {'Nome'}")
    print(f"  {'-'*8} {'-'*40}")
    for item in linhas_sem_nome:
        print(f"  {item['linha']:<8} {item['nome']}")
    print(f"\n  Motivo: {linhas_sem_nome[0]['motivo']}")
    print("\n  Verifique se os nomes estao cadastrados no ServiceNow")
    print("  e corrija a planilha para reprocessar estas linhas.")
else:
    print("Todos os chamados foram abertos com sucesso! Nenhum nome faltou.")

if linhas_sem_numero:
    print(f"\nRELATORIO -- {len(linhas_sem_numero)} chamado(s) aberto(s) mas SEM numero capturado (N/A):\n")
    print(f"  {'Linha':<8} {'Nome'}")
    print(f"  {'-'*8} {'-'*40}")
    for item in linhas_sem_numero:
        print(f"  {item['linha']:<8} {item['nome']}")
    print("\n  O formulario FOI enviado, mas o numero nao apareceu a tempo.")
    print("  Acesse o ServiceNow, filtre por data/hora e corrija a planilha.")
    print("  Na proxima execucao, linhas com N/A serao reprocessadas automaticamente.")

print("=" * 60)
import os
import sys
import time
import re
from datetime import datetime, timedelta

# Força a saída padrão do console para UTF-8 (evita UnicodeEncodeError no Windows/CP1252)
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from src.paths import SOURCE_WORKBOOK, ensure_portable_data_files

# ==============================================================================
# REQUISITOS E REGRAS DE ARQUITETURA
# ==============================================================================
# ATENÇÃO: É ESTRITAMENTE PROIBIDO alterar estes extratores para realizar
# requisições diretas ao servidor (ex: via 'requests', 'urllib', 'curl' ou API).
#
# MOTIVO: Toda a navegação e consulta DEVE ser feita exclusivamente simulando
# o comportamento de tela via Selenium (DOM/Navegador na interface gráfica).
# ==============================================================================


# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

def solicitar_datas_usuario():
    """
    1. Tenta capturar as datas enviadas pela Interface Gráfica via Variáveis de Ambiente.
    2. Se executado avulso via terminal, solicita e valida a digitação das datas.
    3. Retorna no formato YYYY-MM-DD exigido pelos formulários WEB.
    """
    env_ini = os.getenv("GUI_DATA_INICIO")
    env_fim = os.getenv("GUI_DATA_FIM")

    # Integrado com a Interface Gráfica
    if env_ini and env_fim:
        print(f"✅ Datas carregadas da Interface Gráfica: {env_ini} até {env_fim}")
        dt_ini_obj = datetime.strptime(env_ini, "%d/%m/%Y")
        dt_fim_obj = datetime.strptime(env_fim, "%d/%m/%Y")
        return dt_ini_obj.strftime("%Y-%m-%d"), dt_fim_obj.strftime("%Y-%m-%d")

    # Modo interativo (Terminal isolado)
    print("\n" + "=" * 60)
    print("CONFIGURAÇÃO DE DATAS PARA OS EXECUTORES 2 E 3")
    print("=" * 60)
    
    hoje = datetime.now()
    data_padrao_inicio = calcular_dias_uteis_anteriores(hoje, 2).strftime("%d/%m/%Y")
    data_padrao_fim = hoje.strftime("%d/%m/%Y")

    print(f" Sugestão padrão (2 dias úteis atrás até hoje): {data_padrao_inicio} a {data_padrao_fim}")
    print(" Pressione ENTER direto para usar o padrão, ou digite as datas desejadas.\n")

    while True:
        dt_ini_input = input(f"Informe a DATA INICIAL (DD/MM/AAAA) [{data_padrao_inicio}]: ").strip()
        if not dt_ini_input:
            dt_ini_input = data_padrao_inicio
        
        try:
            dt_ini_obj = datetime.strptime(dt_ini_input, "%d/%m/%Y")
            break
        except ValueError:
            print("❌ Formato de data inválido! Use o formato DD/MM/AAAA.")

    while True:
        dt_fim_input = input(f"Informe a DATA FINAL (DD/MM/AAAA) [{data_padrao_fim}]: ").strip()
        if not dt_fim_input:
            dt_fim_input = data_padrao_fim
        
        try:
            dt_fim_obj = datetime.strptime(dt_fim_input, "%d/%m/%Y")
            if dt_fim_obj < dt_ini_obj:
                print("❌ A data final não pode ser menor que a data inicial!")
                continue
            break
        except ValueError:
            print("❌ Formato de data inválido! Use o formato DD/MM/AAAA.")

    dat_inicio_web = dt_ini_obj.strftime("%Y-%m-%d")
    dat_fim_web = dt_fim_obj.strftime("%Y-%m-%d")

    print(f"\n✅ Datas confirmadas para os Executores 2 e 3: {dt_ini_obj.strftime('%d/%m/%Y')} até {dt_fim_obj.strftime('%d/%m/%Y')}\n")
    return dat_inicio_web, dat_fim_web


def iniciar_driver(timeout=300):
    """Cria e inicializa o WebDriver ajustando os timeouts de carregamento de página e scripts."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")

    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.fonts": 2
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)

    # O timeout da conexão HTTP do Selenium é independente dos timeouts do
    # navegador; sem este ajuste, o ChromeDriver pode interromper consultas
    # demoradas após o padrão de 120 segundos.
    driver.command_executor.client_config.timeout = timeout
    driver.set_page_load_timeout(timeout)
    driver.set_script_timeout(timeout)

    wait = WebDriverWait(driver, timeout)
    return driver, wait


def calcular_dias_uteis_anteriores(data_base, dias_uteis=2):
    """Gera a data retrocedendo N dias úteis (desconsiderando Sábado e Domingo)."""
    data_atual = data_base
    contagem = 0
    while contagem < dias_uteis:
        data_atual -= timedelta(days=1)
        if data_atual.weekday() < 5:  # 0 a 4 = Segunda a Sexta
            contagem += 1
    return data_atual


def tentar_converter_data(valor):
    if not valor:
        return valor

    valor = str(valor).strip().replace("às", "").replace("as", "").strip()

    formatos = [
        "%d/%m/%Y", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d-%m-%Y %H:%M"
    ]

    for f in formatos:
        try:
            return datetime.strptime(valor, f)
        except Exception:
            continue

    match = re.search(r"\d{2}/\d{2}/\d{4}( \d{2}:\d{2})?", valor)

    if match:
        try:
            return datetime.strptime(match.group(), "%d/%m/%Y %H:%M")
        except Exception:
            try:
                return datetime.strptime(match.group(), "%d/%m/%Y")
            except Exception:
                pass

    return valor


def tentar_numero(valor):
    if not valor:
        return valor

    valor = str(valor).strip()

    try:
        valor_convertido = valor.replace(".", "").replace(",", ".")
        return float(valor_convertido)
    except Exception:
        return valor


def rgba_para_hex(rgba):
    if not rgba or "rgba(0, 0, 0, 0)" in rgba or "transparent" in rgba:
        return None

    nums = re.findall(r"\d+", rgba)
    if len(nums) >= 3:
        r, g, b = int(nums[0]), int(nums[1]), int(nums[2])
        return f"{r:02X}{g:02X}{b:02X}"

    return None


def salvar_dados_em_aba(caminho_arquivo, nome_aba, dados):
    """
    Salva ou atualiza os dados em uma aba específica no mesmo arquivo Excel.
    """
    if os.path.exists(caminho_arquivo):
        wb = load_workbook(caminho_arquivo)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    if nome_aba in wb.sheetnames:
        del wb[nome_aba]

    ws = wb.create_sheet(title=nome_aba)

    for i, linha in enumerate(dados, start=1):
        for j, celula in enumerate(linha, start=1):
            texto = celula["texto"]
            cor = rgba_para_hex(celula["fundo"])

            valor = tentar_converter_data(texto)

            if not isinstance(valor, datetime):
                valor = tentar_numero(texto)

            cell = ws.cell(row=i, column=j, value=valor)

            if isinstance(valor, datetime):
                cell.number_format = "DD/MM/YYYY"
            else:
                cell.number_format = "General"

            if cor:
                cell.fill = PatternFill(
                    start_color=cor,
                    end_color=cor,
                    fill_type="solid"
                )

    wb.save(caminho_arquivo)
    print(f"Aba '{nome_aba}' salva com sucesso em: {caminho_arquivo}")


def extrair_tabela_dom(driver):
    """
    Função utilitária JS para extração e higienização dos dados da tabela.
    """
    return driver.execute_script("""
        let tabelas = Array.from(document.querySelectorAll('table'));
        if(tabelas.length === 0) return [];

        let maior = tabelas[0];
        for(let t of tabelas){
            if(t.rows.length > maior.rows.length) maior = t;
        }

        let resultado = [];
        for(let r of maior.rows){
            let celulas = Array.from(r.children).filter(el => el.tagName === 'TD' || el.tagName === 'TH');
            
            if(celulas.length === 1 && (celulas[0].colSpan > 3 || celulas[0].getAttribute('colspan') > 3)) {
                continue;
            }

            let linha = [];

            if(celulas.length === 1) {
                let texto = celulas[0].innerText ? celulas[0].innerText.trim() : "";
                let partes = texto.split(/\\t|\\n|  +/);
                
                if(partes.length > 1) {
                    for(let p of partes) {
                        if(p.trim() !== "") {
                            linha.push({
                                texto: p.trim(),
                                fundo: window.getComputedStyle(celulas[0]).backgroundColor
                            });
                        }
                    }
                } else {
                    linha.push({
                        texto: texto,
                        fundo: window.getComputedStyle(celulas[0]).backgroundColor
                    });
                }
            } else {
                for(let c of celulas){
                    let texto = c.innerText ? c.innerText.trim() : "";
                    let estilo = window.getComputedStyle(c);
                    linha.push({
                        texto: texto,
                        fundo: estilo.backgroundColor
                    });
                }
            }

            if(linha.length > 0) {
                resultado.push(linha);
            }
        }
        return resultado;
    """)


def esperar_tabela_estabilizar(driver):
    linhas_anteriores = 0
    repeticoes = 0
    while True:
        linhas = driver.execute_script("""
            let tabela = document.querySelector("table");
            if(!tabela) return 0;
            return tabela.rows.length;
        """)
        if linhas == linhas_anteriores:
            repeticoes += 1
        else:
            repeticoes = 0

        if repeticoes >= 3:
            break
        linhas_anteriores = linhas
        time.sleep(1)


def realizar_login(driver, wait):
    """Acessa e realiza o login com tratamento de espera."""
    driver.get("http://192.168.0.6")
    wait.until(EC.presence_of_element_located((By.NAME, "usuario")))

    user = os.getenv("GUI_USER", "nycolas.santos")
    password = os.getenv("GUI_PASS", "Lider123@")

    driver.find_element(By.NAME, "usuario").send_keys(user)
    driver.find_element(By.NAME, "senha").send_keys(password)

    form = driver.find_element(By.CSS_SELECTOR, "form")
    driver.execute_script("arguments[0].requestSubmit();", form)

    wait.until(
        lambda d: d.find_elements(By.XPATH, "//*[contains(text(),'Olá,')]") or
                  d.find_elements(By.XPATH, "//a[contains(@href,'logoff') or contains(text(),'Sair')]")
    )


def executar_com_tentativas(funcao, *args, **kwargs):
    """Executa uma extração novamente até três vezes quando houver falha."""
    max_tentativas = 3

    for tentativa in range(1, max_tentativas + 1):
        try:
            return funcao(*args, **kwargs)
        except Exception as exc:
            if tentativa == max_tentativas:
                raise

            print(
                f"⚠️ {funcao.__name__} falhou na tentativa "
                f"{tentativa}/{max_tentativas}: {exc}"
            )
            print("🔁 Repetindo somente esta etapa...")


# =====================================================
# EXECUTOR 1: SALDO
# =====================================================

def executar_saldo(arquivo_excel):
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO SALDO (Timeout Estendido: 300s)")
    print("=" * 60)

    driver, wait = iniciar_driver(300)

    try:
        realizar_login(driver, wait)

        intervalos = [
            ("700", "799"),
            ("1000", "1499"),
            ("1500", "1599"),
            ("1600", "3999"),
            ("4000", "4999"),
            ("5000", "5999"),
            ("6000", "6999"),
            ("7000", "7999"),
            ("8000", "8999"),
            ("9000", "9999")
        ]

        abas = []

        for i, (inicio, fim) in enumerate(intervalos):
            if i > 0:
                driver.switch_to.new_window('tab')

            abas.append(driver.current_window_handle)
            driver.get("http://192.168.0.6/pcp/situacao_pedidos2.php")

            wait.until(EC.presence_of_element_located((By.NAME, "cod_empresa")))

            driver.find_element(By.NAME, "cod_empresa").send_keys("11")
            driver.execute_script("document.getElementsByName('inicio')[0].value='2001-01-01'")
            driver.execute_script("document.getElementsByName('fim')[0].value='2050-12-31'")

            driver.find_element(By.NAME, "repi").send_keys(inicio)
            driver.find_element(By.NAME, "repf").send_keys(fim)

            driver.find_element(
                By.CSS_SELECTOR,
                "input[type='radio'][name='show_cod_nat_oper'][value='S']"
            ).click()

            driver.find_element(By.NAME, "exibir_valores").click()
            driver.find_element(By.NAME, "exibir_origem").click()
            driver.find_element(
                By.CSS_SELECTOR,
                "input[type='checkbox'][name='exibir_aviso'][value='sim']"
            ).click()

        todos_dados = []

        for i, aba in enumerate(abas):
            driver.switch_to.window(aba)
            print(f"Executando Saldo {i+1}/{len(intervalos)}")

            try:
                btn_consultar = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'Consultar') or @name='consultar']"))
                )
                btn_consultar.click()
                # Wait for the results table to appear with a longer timeout (up to 600 seconds)
                table_wait = WebDriverWait(driver, 600)
                try:
                    table_wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
                except TimeoutException:
                    # If still not found, attempt a page refresh and retry once
                    print("⚠️ Tabela não encontrada após 600s, recarregando a página e tentando novamente.")
                    driver.refresh()
                    table_wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            except TimeoutException as exc:
                raise RuntimeError(f"Timeout relatório saldo: {exc}")

            esperar_tabela_estabilizar(driver)

            dados = extrair_tabela_dom(driver)

            if i > 0 and len(dados) > 0:
                dados = dados[1:]

            todos_dados.extend(dados)
            print("Linhas acumuladas (Saldo):", len(todos_dados))

        salvar_dados_em_aba(arquivo_excel, "SALDO", todos_dados)

        print("=" * 60)
        print("SALDO FINALIZADO")
        print("=" * 60)

        return todos_dados

    finally:
        driver.quit()


# =====================================================
# EXECUTOR 2: AVISO ENTREGA (VDP)
# =====================================================

def executar_aviso(arquivo_excel, dat_contato, dat_contato_ate):
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO AVISO ENTREGA (VDP)")
    print(f"Período de extração: {dat_contato} até {dat_contato_ate}")
    print("=" * 60)

    driver, wait = iniciar_driver(300)

    try:
        realizar_login(driver, wait)

        intervalos = [
            ("700", "799"),
            ("1000", "9999")
        ]

        abas = []

        for i, (inicio, fim) in enumerate(intervalos):
            if i > 0:
                driver.switch_to.new_window('tab')

            abas.append(driver.current_window_handle)
            driver.get("http://192.168.0.6/vdp/consulta_aviso_entrega2.php")

            wait.until(EC.presence_of_element_located((By.NAME, "cod_empresa")))

            # Empresa
            campo_empresa = driver.find_element(By.NAME, "cod_empresa")
            campo_empresa.clear()
            campo_empresa.send_keys("11")

            # Datas customizadas fornecidas
            driver.execute_script(f"document.getElementsByName('dat_contato')[0].value='{dat_contato}';")
            driver.execute_script(f"document.getElementsByName('dat_contato_ate')[0].value='{dat_contato_ate}';")

            # Representantes
            campo_inicio = driver.find_element(By.NAME, "inicio")
            campo_inicio.clear()
            campo_inicio.send_keys(inicio)

            campo_fim = driver.find_element(By.NAME, "fim")
            campo_fim.clear()
            campo_fim.send_keys(fim)

            # Checkboxes de Tipos
            for val in ["AE", "OE", "AD"]:
                chk = driver.find_elements(By.XPATH, f"//input[@name='tipos[]'][@value='{val}']")
                if chk and not chk[0].is_selected():
                    chk[0].click()

            # Apenas Pedidos em Aberto
            chk_abertos = driver.find_elements(By.XPATH, "//input[@name='pedidos_abertos'][@value='S']")
            if chk_abertos and not chk_abertos[0].is_selected():
                chk_abertos[0].click()

        todos_dados = []

        for i, aba in enumerate(abas):
            driver.switch_to.window(aba)
            print(f"Executando Aviso Entrega ({intervalos[i][0]}-{intervalos[i][1]})")

            try:
                btn_consultar = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'consultar') or contains(@value,'Consultar')]"))
                )
                btn_consultar.click()
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            except TimeoutException as exc:
                raise RuntimeError(f"Timeout relatório aviso entrega: {exc}")

            esperar_tabela_estabilizar(driver)

            dados = extrair_tabela_dom(driver)

            if i > 0 and len(dados) > 0:
                dados = dados[1:]

            todos_dados.extend(dados)
            print("Linhas acumuladas (Aviso Entrega):", len(todos_dados))

        salvar_dados_em_aba(arquivo_excel, "AVISO_ENTREGA", todos_dados)

        print("=" * 60)
        print("AVISO ENTREGA FINALIZADO")
        print("=" * 60)

        return todos_dados

    finally:
        driver.quit()


# =====================================================
# EXECUTOR 3: CONSULTA PEDIDOS (VDP)
# =====================================================

def executar_consulta_pedidos(arquivo_excel, dat_emis, dat_emis1):
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO CONSULTA PEDIDOS (VDP)")
    print(f"Período de extração: {dat_emis} até {dat_emis1}")
    print("=" * 60)

    driver, wait = iniciar_driver(300)

    try:
        realizar_login(driver, wait)

        intervalos_repres = [
            ("700", "799"),
            ("1000", "9999")
        ]

        abas = []

        for i, (cod_repres, cod_repres1) in enumerate(intervalos_repres):
            if i > 0:
                driver.switch_to.new_window('tab')

            abas.append(driver.current_window_handle)
            driver.get("http://192.168.0.6/vdp/consulta_pedidos.php")

            wait.until(EC.presence_of_element_located((By.NAME, "cod_empresa")))

            driver.find_element(By.NAME, "cod_empresa").send_keys("11")
            
            driver.find_element(By.NAME, "cod_repres").clear()
            driver.find_element(By.NAME, "cod_repres").send_keys(cod_repres)

            driver.find_element(By.NAME, "cod_repres1").clear()
            driver.find_element(By.NAME, "cod_repres1").send_keys(cod_repres1)

            # Datas customizadas fornecidas
            driver.execute_script(f"document.getElementsByName('dat_emis')[0].value='{dat_emis}';")
            driver.execute_script(f"document.getElementsByName('dat_emis1')[0].value='{dat_emis1}';")

        todos_dados = []

        for i, aba in enumerate(abas):
            driver.switch_to.window(aba)
            print(f"Executando Pedidos ({intervalos_repres[i][0]}-{intervalos_repres[i][1]})")

            try:
                btn_consultar = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'Consultar') or @name='consultar']"))
                )
                btn_consultar.click()
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            except TimeoutException as exc:
                raise RuntimeError(f"Timeout relatório consulta pedidos: {exc}")

            esperar_tabela_estabilizar(driver)

            dados = extrair_tabela_dom(driver)

            if i > 0 and len(dados) > 0:
                dados = dados[1:]

            todos_dados.extend(dados)
            print("Linhas acumuladas (Pedidos):", len(todos_dados))

        salvar_dados_em_aba(arquivo_excel, "PEDIDOS", todos_dados)

        print("=" * 60)
        print("CONSULTA PEDIDOS FINALIZADA")
        print("=" * 60)

        return todos_dados

    finally:
        driver.quit()


# =====================================================
# EXECUTOR 4: ESTOQUE LOTE ENDEREÇADO (SUP)
# =====================================================

def executar_estoque_lote(arquivo_excel):
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO ESTOQUE LOTE ENDEREÇADO (SUP)")
    print("=" * 60)

    driver, wait = iniciar_driver(300)

    try:
        realizar_login(driver, wait)

        driver.get("http://192.168.0.6/sup/estoque_lote_ender.php")
        wait.until(EC.presence_of_element_located((By.NAME, "cod_empresa")))

        campo_empresa = driver.find_element(By.NAME, "cod_empresa")
        campo_empresa.clear()
        campo_empresa.send_keys("11")

        radio_situacao = driver.find_elements(By.XPATH, "//input[@type='radio'][@name='situacao'][@value='L']")
        if radio_situacao:
            radio_situacao[0].click()

        campo_endereco = driver.find_element(By.NAME, "endereco")
        campo_endereco.clear()
        campo_endereco.send_keys("VT")

        radio_tipo = driver.find_elements(By.XPATH, "//input[@type='radio'][@name='tipo'][@value='F']")
        if radio_tipo:
            radio_tipo[0].click()

        radio_relatorio = driver.find_elements(By.XPATH, "//input[@type='radio'][@name='relatorio'][@value='R']")
        if radio_relatorio:
            radio_relatorio[0].click()

        btn_consultar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'Consultar') or @type='submit']"))
        )
        btn_consultar.click()

        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        esperar_tabela_estabilizar(driver)

        dados = extrair_tabela_dom(driver)

        salvar_dados_em_aba(arquivo_excel, "ESTOQUE_LOTE", dados)

        print("=" * 60)
        print("ESTOQUE LOTE ENDEREÇADO FINALIZADO")
        print("=" * 60)

        return dados

    finally:
        driver.quit()


# =====================================================
# PONTO DE ENTRADA PRINCIPAL
# =====================================================

if __name__ == "__main__":
    ensure_portable_data_files()
    ARQUIVO_CONSOLIDADO = str(SOURCE_WORKBOOK)

    # 1. Carrega/solicita as datas necessárias para os Executores 2 e 3
    dt_inicio_e2_e3, dt_fim_e2_e3 = solicitar_datas_usuario()

    # 2. Execução sequencial das rotinas
    executar_com_tentativas(executar_saldo, ARQUIVO_CONSOLIDADO)
    executar_com_tentativas(
        executar_aviso,
        ARQUIVO_CONSOLIDADO,
        dt_inicio_e2_e3,
        dt_fim_e2_e3,
    )
    executar_com_tentativas(
        executar_consulta_pedidos,
        ARQUIVO_CONSOLIDADO,
        dt_inicio_e2_e3,
        dt_fim_e2_e3,
    )
    executar_com_tentativas(executar_estoque_lote, ARQUIVO_CONSOLIDADO)

    print("\n" + "=" * 60)
    print("TODAS AS EXTRAÇÕES FORAM CONCLUÍDAS COM SUCESSO!")
    print(f"Arquivo final unificado salvo em: {ARQUIVO_CONSOLIDADO}")
    print("=" * 60)

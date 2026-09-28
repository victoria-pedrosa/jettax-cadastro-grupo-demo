import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

def extrair_cnpj(texto):
    """Extrai apenas o CNPJ da célula."""
    padrao_cnpj = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', str(texto))
    if padrao_cnpj:
        return padrao_cnpj.group()
    return str(texto).strip()

# ---------------------------------------------------------
# 1. LEITURA DA PLANILHA
# ---------------------------------------------------------
print("Lendo a planilha de dados...")
try:
    df = pd.read_excel("dados_grupos.xlsx")
except Exception as e:
    print(f"Erro ao ler a planilha: {e}")
    exit()

resultados = []

# ---------------------------------------------------------
# 2. INICIALIZAÇÃO DO NAVEGADOR
# ---------------------------------------------------------
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 10)

url_cadastro = "https://admin.jettax360.com.br/client/group/form"
driver.get(url_cadastro)

print("\n" + "="*50)
input("Faça o login, navegue até a tela 'Cadastrar grupo' e pressione ENTER aqui no terminal para iniciar...")
print("="*50 + "\n")

# ---------------------------------------------------------
# 3. EXECUÇÃO DA AUTOMAÇÃO
# ---------------------------------------------------------
for nome_grupo in df.columns:
    linhas_coluna = df[nome_grupo].dropna().tolist()

    if not linhas_coluna:
        continue

    print(f"\n--- Iniciando grupo: {nome_grupo} ({len(linhas_coluna)} clientes) ---")

    try:
        # Preencher o nome do grupo
        campo_descricao = wait.until(EC.presence_of_element_located((By.NAME, "name")))
        campo_descricao.clear()
        campo_descricao.send_keys(str(nome_grupo))

        for linha in linhas_coluna:
            cnpj = extrair_cnpj(linha)
            print(f"Buscando: {cnpj}")

            try:
                # Limpa e busca
                campo_busca = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Razão Social ou CNPJ']")))
                campo_busca.send_keys(Keys.CONTROL + "a")
                campo_busca.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)

                campo_busca.send_keys(cnpj)
                time.sleep(2)  # Aguarda o sistema filtrar

                # --- LOCALIZA O CHECKBOX REAL DA LINHA FILTRADA ---
                # A tabela usa vue-recycle-scroller: linhas antigas continuam no DOM
                # fora da tela. Por isso localizamos pelo TEXTO do CNPJ dentro do
                # item-view (evita pegar um nó reciclado/invisível).
                xpath_linha = (
                    "//div[contains(@class,'vue-recycle-scroller__item-view') and "
                    f"contains(., '{cnpj}')]//input[@type='checkbox']"
                )
                checkbox_input = wait.until(EC.presence_of_element_located((By.XPATH, xpath_linha)))

                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox_input)
                time.sleep(0.3)

                # --- MARCA VIA TECLADO (não via clique) ---
                # O <label> do switch tem tamanho 0x0 e intercepta o hit-test do
                # clique de forma inconsistente. Foco no <input> + tecla ESPAÇO é o
                # jeito nativo do navegador de marcar um checkbox e funciona sempre,
                # testado e confirmado ao vivo no sistema.
                checkbox_input.send_keys(Keys.SPACE)
                time.sleep(0.3)

                if checkbox_input.is_selected():
                    resultados.append({"Grupo": nome_grupo, "CNPJ/Linha": linha, "CNPJ Buscado": cnpj, "Status": "Vinculado com sucesso"})
                    print(" -> Vinculado!")
                else:
                    raise Exception("Espaço não marcou o input.")

            except Exception as e:
                print(f" -> ERRO: {e}")
                resultados.append({"Grupo": nome_grupo, "CNPJ/Linha": linha, "CNPJ Buscado": cnpj, "Status": f"Erro: {e}"})

            # Limpa o campo para a próxima busca
            campo_busca.send_keys(Keys.CONTROL + "a")
            campo_busca.send_keys(Keys.BACKSPACE)
            time.sleep(1)

        # Salvar o grupo
        print("Salvando grupo...")
        try:
            botao_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Salvar grupo')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_salvar)
            time.sleep(0.3)
            botao_salvar.click()
            print(f"[+] Grupo '{nome_grupo}' salvo com sucesso!")
        except Exception as e:
            print(f"[-] Erro ao salvar: {e}")

        # Aguarda 1 minuto e recarrega a página (F5) antes do próximo grupo
        print("Aguardando 1 minuto antes do próximo grupo...")
        time.sleep(60)
        driver.refresh()
        wait.until(EC.presence_of_element_located((By.NAME, "name")))
        time.sleep(2)

    except Exception as e:
        print(f"Erro fatal ao processar o grupo {nome_grupo}. Pulando para o próximo. Erro: {e}")
        driver.get(url_cadastro)

# ---------------------------------------------------------
# 4. GERAÇÃO DA PLANILHA DE RESULTADOS
# ---------------------------------------------------------
print("\nGerando planilha de resultados...")
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel("resultado_processamento.xlsx", index=False)

print("Automação finalizada. Verifique o arquivo 'resultado_processamento.xlsx'.")
driver.quit()
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from colorama import Fore, Style, init
import time

init(autoreset=True) # Inicializa com autoreset para evitar restauração manual de estilos

def contador_iteração(inicio, fim):
    tempo_duracao = fim - inicio
    print(f'\n\tEssa Iteração durou {Fore.YELLOW}{tempo_duracao:.2f}{Style.RESET_ALL} SEGUNDOS')
def contador_tempototal(inicio_total, fim):
    tempo_duracao = fim - inicio_total
    horas, resto = divmod(tempo_duracao, 3600) 
    minutos, segundos = divmod(resto, 60) 
    tempo_formatado = f"{int(horas):02}:{int(minutos):02}:{int(segundos):02}" # Formata como HH:MM:SS
    print(f"\t\tTempo percorrido até o momento: {Fore.YELLOW}{tempo_formatado}")

def save_excel():                    # Função para salvar todos os DataFrames em um único arquivo Excel, com respectivas abas
    if not os.path.exists("XLS's"):         # Se o excel não exister, ele cria um novo
        os.makedirs("XLS's")
    with pd.ExcelWriter("XLS's/Contratos 12000 a 14499 EXCEL.xlsx", engine='openpyxl') as writer:
        df_new.to_excel(writer, sheet_name='Contratos', index=False)
        df_entidades.to_excel(writer, sheet_name='Entidades Vinculadas', index=False)
        df_inteiro_teor.to_excel(writer, sheet_name='Termos Vinculados', index=False)
        df_empenhos.to_excel(writer, sheet_name='Empenhos', index=False)
        df_error.to_excel(writer, sheet_name='Erros', index=False)
    print("\tExcel atualizado salvo com sucesso!")

inicio_total = time.time()
edge_driver_path = 'msedgedriver.exe' # Caminho completo para o EdgeDriver

# Carregamento do CSV com contratos para extração do campo 'Nº Instrumento'
df = pd.read_excel("CSV's\Linhas 12000 a 14499 CSV_.xlsx")  # Lê o xlsx original
contratos = df['Nº Instrumento'].tolist()

# Criação de DataFrame's 
df_new = pd.DataFrame(columns=[
    'Número de Instrumento', 'Tipo de Instrumento', 'Data de Publicação', 'Situação',
    'Período Inicial', 'Período Final', 'Valor Total (com aditivos)', 'Entidades Vinculadas',
    'Termos Vinculados', 'Empenhos Emitidos', 'Valor Total de empenhos', 'Objeto'
])
df_entidades = pd.DataFrame(columns=['Número de Instrumento', 'Entidade', 'CNPJ'])
df_inteiro_teor = pd.DataFrame(columns=['Número de Instrumento', 'Termo'])
df_empenhos = pd.DataFrame(columns=['Número de Instrumento', 'Número de Empenho', 'Valor Empenho', 'Descrição Empenho'])
df_error = pd.DataFrame(columns=['Número de Instrumento'])

# Inicialização do WebDriver para o Microsoft Edge
service = Service(executable_path=edge_driver_path)
driver = webdriver.Edge(service=service)
wait = WebDriverWait(driver, 5) # Definição de espera
waitUrl = WebDriverWait(driver, 25)

url = "https://www.codevasf.gov.br/acesso-a-informacao/licitacoes-e-contratos/contratos" # Acessar o site
driver.get(url)

contador_total = 0
contador_refresh = 0
# Loop pelos contratos
for i, contrato in enumerate(contratos):
    inicio = time.time()
    sucesso = False
    tentativas = 0
    entidades = []
    inteiro_teor = []
    empenhos = []
    print(f"\n\nIniciando processamento para o contrato: {Fore.BLUE}{Style.BRIGHT}{contrato}")

    while not sucesso and tentativas < 2:  # Efetuar até 3 tentativas
        try:
            contador_refresh += 1
            contador_total += 1
            tentativas += 1
            print(f"\tTentativa {tentativas} para o contrato {contrato}")

            waitUrl.until(EC.presence_of_element_located((By.XPATH,'//*[@id="exercicio"]/option[1]'))) #Espera a opção (Todos estar presente no campo)
            waitUrl.until(EC.presence_of_element_located((By.XPATH,'//*[@id="uf"]/option[1]'))) #Espera a opção (Todos estar presente no campo)

            campo_contrato = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="numeroInstrumento"]'))) # Preencher o campo de contrato
            driver.execute_script("arguments[0].scrollIntoView();", campo_contrato)
            campo_contrato.clear()
            campo_contrato.send_keys(contrato) # print(f"Campo contrato preenchido com {contrato}")

            botao_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnPesquisar"]')))  # Clicar no botão de pesquisa
            botao_pesquisar.click() # print("Botão de pesquisa clicado.")

            resultlist = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="resultList"]/div'))).text
            botao_listar = driver.find_element(By.XPATH, '//*[@id="btnListar"]')
            if " Nenhum contrato atende aos critérios de pesquisa." in resultlist:
                botao_listar.click()
            else:
                wait.until(EC.element_to_be_clickable(botao_listar)).click()
            
            for linha in range(4, 9):  # Loop nas linhas 4 a 8
                try:
                    xpath_contrato = f'//*[@id="quadroContratos"]/div/table/tbody/tr[{linha}]/td[1]/a'
                    valor_pesquisado = wait.until(EC.presence_of_element_located((By.XPATH, xpath_contrato))).text

                    if valor_pesquisado == contrato:
                        print(f"\tContrato encontrado: {contrato} na linha {linha}")
                        js_link = driver.find_element(By.XPATH, xpath_contrato)
                        js_link.click()
                        break
                except NoSuchElementException:
                    print(f"\n\tContrato {contrato} não encontrado na linha {linha}, tentando próxima linha.")
                    pass

            # Extração de elementos
            elemento = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="modalPanel"]/div/table/tbody/tr[2]/td'))).text
            tipo_instrumento, numero = elemento.rsplit(' ', 1) # print(f"Tipo de Instrumento: {tipo_instrumento}")

            data_publicacao = wait.until(EC.visibility_of_element_located((By.XPATH, '//td[text()="Data de Publicação :"]/following-sibling::td'))).text
            if data_publicacao == '':
                data_publicacao = 'Não existe' #print(f"Data de Publicação: {data_publicacao}")

            periodo_vigencia = wait.until(EC.visibility_of_element_located((By.XPATH, '//td[text()="Período de Vigência :"]/following-sibling::td'))).text
            if ' - ' in periodo_vigencia:
                periodo_inicial, periodo_final = periodo_vigencia.split(' - ')
            elif periodo_vigencia == '-':
                periodo_inicial, periodo_final = 'Não existe', 'Não existe' #print(f"Período de Vigência: {periodo_inicial} - {periodo_final}")

            try:
                objeto = driver.find_element(By.XPATH, '//td[text()="Objeto :"]/following-sibling::td').text # print(f"Objeto: {objeto}")
            except NoSuchElementException:
                objeto = 'Não Disponível' #print("Objeto não encontrado.")

            try:
                situacao = wait.until(EC.visibility_of_element_located((By.XPATH, '//td[text()="Situação:"]/following-sibling::td'))).text # print(f"Situação: {situacao}")
            except (NoSuchElementException, TimeoutException):
                situacao = 'Não Disponível' # print("Situação não encontrada.")

            try:
                valor_total = wait.until(EC.visibility_of_element_located((By.XPATH, '//td[text()="Valor Total (com aditivos) :"]/following-sibling::td'))).text # print(f"Valor Total (com aditivos): {valor_total}")
            except (NoSuchElementException, TimeoutException):
                valor_total = 'Não Disponível' # print("Valor Total (com aditivos) não encontrado.")

            try:
                tr_entidades = wait.until(EC.presence_of_element_located((By.XPATH, '//td[text()="Entidades Vinculadas :"]/ancestor::tr')))
                tr_seguinte_entidades = tr_entidades.find_elements(By.XPATH, 'following-sibling::tr')
                
                for tr in tr_seguinte_entidades:
                    primeiro_td_entidades = tr.find_element(By.XPATH, './td[1]').text.strip()
                    if primeiro_td_entidades == '':
                        cnpj = tr.find_element(By.XPATH, './td[2]').text.strip()
                        nome_entidade = tr.find_element(By.XPATH, './td[3]').text.strip()
                        entidades.append({'CNPJ': cnpj, 'Nome da Entidade': nome_entidade})
                    else:
                        break

                qtd_entidades = len(entidades) # print(f"Quantidade de Entidades Vinculadas: {qtd_entidades}") # print(f"Entidades Capturadas:")
                for entidade in entidades: # print(f"\t{entidade['CNPJ']}, {entidade['Nome da Entidade']}")
                    df_entidades = pd.concat([df_entidades, pd.DataFrame([{ 'Número de Instrumento': contrato, 
                        'Entidade': entidade['Nome da Entidade'], 'CNPJ': entidade['CNPJ'] }])], ignore_index=True)
            except (NoSuchElementException, TimeoutException):
                qtd_entidades = 0
                pass
            
            try:
                tr_inteiro_teor = driver.find_element(By.XPATH, '//td[text()="Inteiro Teor :"]/ancestor::tr')
                tr_seguinte_inteiro_teor = tr_inteiro_teor.find_elements(By.XPATH, 'following-sibling::tr')
                
                for tr in tr_seguinte_inteiro_teor:
                    primeiro_td_inteiro_teor = tr.find_element(By.XPATH, './td[1]').text.strip()
                    if primeiro_td_inteiro_teor == '':
                        termo = tr.find_element(By.XPATH, './td[2]').text.strip()
                        inteiro_teor.append(termo)
                    else:
                        break

                qtd_inteiro_teor = len(inteiro_teor) # print(f"Quantidade de Termos Vinculados: {qtd_inteiro_teor}") # print(f"Termos Capturados:")
                for termo in inteiro_teor: # print(f'\t{termo}')
                    df_inteiro_teor = pd.concat([df_inteiro_teor, pd.DataFrame([{
                        'Número de Instrumento': contrato, 'Termo': termo }])], ignore_index=True)
            except (NoSuchElementException, TimeoutException):
                qtd_inteiro_teor = 0
                pass

            try:
                tr_empenhos = wait.until(EC.presence_of_element_located((By.XPATH, '//td[text()="Empenhos Emitidos :"]/ancestor::tr')))
                tr_seguinte_empenhos = tr_empenhos.find_elements(By.XPATH, 'following-sibling::tr')
                for tr in tr_seguinte_empenhos:
                    try:
                        numero_empenho_element = tr.find_element(By.XPATH, './td[2]/a')
                        numero_empenho = numero_empenho_element.text.strip()
                        link_empenho = numero_empenho_element.get_attribute('href')
                        valor_empenho = tr.find_element(By.XPATH, './td[3]').text.strip()
                        valor_empenho = valor_empenho.replace('.', '').replace(',', '.')
                        descricao_empenho = tr.find_element(By.XPATH, './td[4]').text.strip()
                        empenhos.append({ 'Número de Empenho': numero_empenho, 'Link de Empenho': link_empenho,
                            'Valor Empenho': valor_empenho, 'Descrição Empenho': descricao_empenho })
                    except NoSuchElementException:
                        break #Quando não existir mais empenhos na lista
                    
                qtd_empenhos = len(empenhos) 
                valor_total_empenhos = sum(float(empenho['Valor Empenho']) for empenho in empenhos) # print(f"Quantidade de Empenhos Emitidos: {qtd_empenhos}") # print(f"Valor Total de Empenhos: {valor_total_empenhos}") # print(f"Empenhos Capturados:")
                for empenho in empenhos: # print(f"\t{empenho['Número de Empenho']}, {empenho['Valor Empenho']}")
                    df_empenhos = pd.concat([df_empenhos, pd.DataFrame([{ 'Número de Instrumento': contrato,
                        'Número de Empenho': f'=HYPERLINK("{empenho["Link de Empenho"]}", "{empenho["Número de Empenho"]}")',
                        'Valor Empenho': empenho['Valor Empenho'], 'Descrição Empenho': empenho['Descrição Empenho'] }])], ignore_index=True)

            except (NoSuchElementException, TimeoutException):
                valor_total_empenhos = 0
                qtd_empenhos = 0
                pass

            # Adiciona dados ao DataFrame Principal
            df_new = pd.concat([df_new, pd.DataFrame([{ 'Número de Instrumento': contrato,'Tipo de Instrumento': tipo_instrumento, 'Data de Publicação': data_publicacao, 
                'Situação': situacao, 'Período Inicial': periodo_inicial, 'Período Final': periodo_final, 'Valor Total (com aditivos)': valor_total, 'Entidades Vinculadas': qtd_entidades, 
                'Termos Vinculados': qtd_inteiro_teor, 'Empenhos Emitidos': qtd_empenhos, 'Valor Total de empenhos': valor_total_empenhos,'Objeto': objeto }])], ignore_index=True)
            
            save_excel()# Salvar tudo no excel
            sucesso = True  # Se tudo correr bem, a pesquisa foi um sucesso
            
            botao_fechar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="closeModal"]')))
            botao_fechar.click()# Fechar o POP-UP com o contrato

            botao_listarnovo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnVoltar"]')))
            botao_listarnovo.click()# Clicar no botão de Nova Consulta

            fim = time.time()
            contador_iteração(inicio = inicio,fim = fim)
            print(f"\n\tN° de ITERAÇÕES para refresh: {Fore.RED}{Style.BRIGHT}{contador_refresh}")
            print(f"\t\tN° de contratos TOTAIS: {Fore.LIGHTMAGENTA_EX}{contador_total}")
            contador_tempototal(inicio_total = inicio_total, fim = fim)

            if contador_refresh >= 249:
                driver.refresh()
                contador_refresh = 0  # Reinicia o contador após o refresh
                waitUrl.until(EC.presence_of_element_located((By.XPATH, '//*[@id="exercicio"]/option[1]')))
                waitUrl.until(EC.presence_of_element_located((By.XPATH, '//*[@id="uf"]/option[1]')))
                print(f"\n\t{Fore.LIGHTBLUE_EX}Site recarregado com sucesso após refresh.")

        except NoSuchElementException:
            print('Botão nao encontrado')
        except TimeoutException:
            driver.get(url)
            print(f"\n\t{Fore.LIGHTBLUE_EX}Site acessado com sucesso após driver.get()")
            
        except ElementNotInteractableException:
            print(f"\t{Fore.RED}{Style.BRIGHT}Erro ao processar o contrato: {contrato}. Tentativa {tentativas}.")
            botao_listarnovo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnVoltar"]')))
            botao_listarnovo.click()# Clicar no botão de Nova Consulta
            if tentativas == 2:
                print(f"\tFalha ao processar o contrato {contrato} após 2 tentativas.")
                df_error = pd.concat([df_error, pd.DataFrame([{'Número de Instrumento': contrato}])], ignore_index=True)
                save_excel()
                print(f'\tContrato: {contrato} salvo em lista error')
                fim = time.time()
                contador_iteração(inicio = inicio,fim = fim)
                contador_total -= 1
                print(f"\t\tN° de ITERAÇÕES realizadas: {Fore.RED}{contador_refresh-1}")
                print(f"\t\tN° de contratos TOTAIS: {Fore.LIGHTMAGENTA_EX}{contador_total}")
                contador_tempototal(inicio_total = inicio_total, fim = fim)
                break

driver.quit() # Fechar o navegador
print(f'{Fore.LIGHTGREEN_EX}{Style.BRIGHT}Processo concluído. Todos os contratos foram processados!!!')
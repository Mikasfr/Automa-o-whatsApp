from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from urllib.parse import quote
import time
import random
from datetime import datetime

# Caminho para o ChromeDriver
chrome_driver_path = '/home/mk/WorkSpace/chromedriver-linux64/chromedriver'

# Função para verificar se o horário está entre 23:00 e 08:00
def dentro_do_horario_proibido():
    hora_atual = datetime.now().hour
    return hora_atual >= 23 or hora_atual < 8

# Função para pausar por um intervalo aleatório entre 10 e 20 minutos
def pausa_aleatoria():
    minutos_pausa = random.randint(10, 20)
    print(f'Pausando por {minutos_pausa} minutos...')
    time.sleep(minutos_pausa * 60)

# Função para formatar o número de telefone
def formatar_numero_telefone(telefone):
    telefone = str(telefone)  # Garante que o telefone seja uma string
    if not telefone.startswith('+') and telefone != "None":  # Ignora valores vazios
        telefone = '+55' + telefone  # Adiciona o código do Brasil se não estiver presente
    return telefone

# Configura o driver do Chrome
service = ChromeService(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)
driver.get('https://web.whatsapp.com/')

print("Por favor, escaneie o QR code e faça login no WhatsApp Web.")
input("Pressione Enter após escanear o QR code e o WhatsApp Web estar pronto...")

# Ler planilha e guardar informações sobre nome e telefone
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['cliente']

# Variável para controlar o tempo de envio contínuo
inicio_envio = time.time()

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    
    # Se o telefone estiver em branco ou None, pula para a próxima linha
    if not telefone:
        print(f'Número de telefone vazio para {nome}. Pulando...')
        continue
    
    # Verifica se está dentro do horário proibido (23:00 às 08:00)
    while dentro_do_horario_proibido():
        print("Horário proibido. Aguardando até as 08:00...")
        time.sleep(60 * 60)  # Pausa por uma hora e verifica novamente
    
    # Se passou 1 hora desde o início do envio, faça uma pausa aleatória de 10 a 20 minutos
    if time.time() - inicio_envio >= 3600:
        pausa_aleatoria()
        inicio_envio = time.time()  # Reinicia o contador de tempo após a pausa
    
    # Formata o número de telefone para incluir o código do Brasil
    telefone_formatado = formatar_numero_telefone(telefone)

    mensagem = (f'Olá, {nome}\n'
                'Estamos com uma nova plataforma de pedidos com fotos reais das pizzas! No seu primeiro pedido, use o cupom EXCLUSIVO10 e ganhe 10% de desconto, Corre que esse desconto é válido até o dia 30.\n\n'
                'Faça seu pedido em: http://pedir.to/finnadelivery\n\n'
                'Precisa de ajuda? Só me responder aqui que te envio um tutorial.\n\n'
                'Abraço,\n'
                'Equipe da Finna.')

    
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone_formatado}&text={quote(mensagem)}'
    
    driver.get(link_mensagem_whatsapp)
    
    try:
        # Espera a página carregar completamente
        time.sleep(5)  # Pausa para garantir que a página carregue

        # Espera até o campo de mensagem estar presente
        print("Esperando o campo de mensagem...")
        campo_mensagem = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p'))
        )
        print("Campo de mensagem encontrado.")

        # Clica no campo de mensagem para garantir que ele esteja ativo
        campo_mensagem.click()
        print("Campo de mensagem ativado.")

        # Envia a mensagem
        campo_mensagem.clear()  # Limpa o campo antes de digitar

        # Espera o botão de envio estar presente e clicável
        print("Esperando o botão de envio...")
        enviar_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button/span'))
        )
        print("Botão de envio encontrado e clicado.")
        enviar_btn.click()
        
        # Aguarda a confirmação de envio
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@data-testid='msg-time']"))
        )

        print(f'Mensagem enviada para {nome}')
        time.sleep(1)  # Pequena pausa para evitar problemas com envio rápido
        
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}. Erro: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone},{e}\n')

# Encerra o driver
driver.quit()


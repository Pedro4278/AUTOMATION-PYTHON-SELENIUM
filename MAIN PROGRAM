import ctypes
import os
import keyboard
import re
import time
from datetime import datetime
from PIL import Image
import phonenumbers
import pywhatkit
import PySimpleGUI as sg
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import requests
import copy

# Prevent system from going into sleep mode
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
texto_cadastrado=0

def iniciar_driver():
    chrome_options=Options()
    # Work with minimized window
    # chrome_options.add_argument("--headless")

    # Path to Google Chrome user data directory # Starts Chrome with saved cookies!!
    chrome_options.add_argument("user-data-dir=C:\\Users\\vilma\\AppData\\Local\\Google\\Chrome\\User Data\\Default") # Don't forget to include 'user-data-dir' at the beginning

    driver_path = 'C:\\Users\\vilma\\OneDrive\\Área de Trabalho\\webdriver\\chromedriver' # Replace with correct chromedriver.exe path

    service = webdriver.chrome.service.Service(driver_path)
    service.start()

    arguments=['--lang=ptBR','--start-maximize']
    for argument in arguments:
        chrome_options.add_experimental_option('prefs',{
            'download_prompt_for_download':False,
            'profile.default_content_setting_values.notifications':2,
            'profile.default_content_setting_values.automatic_dowloads':1
        })

    driver=webdriver.Chrome(service = ChromeService(ChromeDriverManager().install()),options=chrome_options)
    return driver

def espere_e_clique(driver, by, locator, timeout=15):
    # Wait and click element
    wait = WebDriverWait(driver, timeout)
    button = wait.until(EC.element_to_be_clickable((by, locator)))
    button.click()

def aguardar_carregar(driver,by,locator,timeout=30):
    # Wait for element to load
    aguarde=WebDriverWait(driver,timeout)
    até=aguarde.until(EC.element_to_be_clickable((by, locator)))

def fazer_requisicao(url):
    # Continuously make requests
    while True:
        try:
            response = requests.get(url) # Make GET request

            if response.status_code == 200:
                print('Request successful')
                break  # Exit loop
        except requests.exceptions.ConnectionError:
            print('Connection failed, waiting 120 seconds')
            time.sleep(120)

def final(texto):
    # Extract phone numbers from text
    contatos = []
    for match in phonenumbers.PhoneNumberMatcher(texto, 'BR'):
        numero = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
        contatos.append(numero)
    return contatos

url='https://www.google.com.br/'

# GUI Theme Setup
sg.theme('DarkBlue3')
cor_botao=('white','#FF6600')
cor_de_fundo='#000000'
fonte_geral_botões='Montserrat',15,'bold'
sg.theme_background_color(cor_de_fundo)
imagem_caminho='C:/Users/vilma/OneDrive/Área de Trabalho/logo_jaguar.py.png'
imagem2=Image.open(imagem_caminho)
imagem_corrigida=imagem2.resize((300,200))
imagem_corrigidapy=imagem_corrigida.tobytes()
contatos = []

# Main Window Layout
layout_principal=[
    [sg.Image(r'C:/Users/vilma/OneDrive/Área de Trabalho/logo_jaguar.py.png', size=(350,150),background_color='#000000')],
    [sg.Text('How can I help you today?', size=(20,3), font=('Roboto',25,'bold'),background_color='#000000')],
    [sg.Checkbox('Run in foreground', font=('Roboto',10,'bold'),background_color='#000000', key='modo_headless')],
    [sg.Button('Set Texts', size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
    [sg.Button('Wait for Contacts', size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
    [sg.Button('Send Messages',size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
    [sg.Button('Return to Clients',size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
    [sg.Button('Dismiss Contacts',key='dispensar',size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
    [sg.Button('Login',key='login',size=(25,2), button_color=cor_botao, font=fonte_geral_botões,pad=((50, 50), 0))],
]

janela_principal=sg.Window('Menu',layout_principal)

# Main window loop
senha=0
while senha==0:
  try:
    while True:
       evento_principal,valores_principal = janela_principal.read()

       if evento_principal == sg.WIN_CLOSED:
            janela_principal.close()
            senha=1
            break

       # Set texts to be sent
       if evento_principal == 'Definir Textos' :
           janela_principal.hide()

           layout_texto = [
               [sg.Text('Enter the message you want to send to\nyour clients',key='titulo_texto', size=(40, 2),
                        font=('Roboto', 18, 'bold'), background_color='#000000')],
               [sg.Text('1st Message (required field)', size=(35, 2), font=('Roboto', 14, ), background_color='#000000')],
               [sg.Multiline(key='texto', size=(50, 4), font=('Roboto', 12))],
               [sg.Text('2nd Message (Press space to leave blank)', size=(35, 2), font=('Roboto', 14,), background_color='#000000')],
               [sg.Multiline(key='texto1', size=(50, 4), font=('Roboto', 12))],
               [sg.Text('3rd Message (Press space to leave blank)', size=(35, 2), font=('Roboto', 14, ), background_color='#000000')],
               [sg.Multiline(key='texto2', size=(50, 4), font=('Roboto', 12))],
               [sg.Button('OK', button_color=cor_botao, font=fonte_geral_botões, key='ok_texto'),
                sg.Button('Back', button_color=cor_botao, font=fonte_geral_botões, key='cancelar_texto')],
           ]

           janela_texto = sg.Window('Register Text', layout_texto, size=(580, 570))

           while True:
            evento_texto, valores_texto = janela_texto.read()

            if evento_texto == sg.WIN_CLOSED:
                janela_texto.close()
                janela_principal.un_hide()
                break

            if evento_texto == 'ok_texto' and valores_texto['texto'] != '':
               # Variable that stores texts to be sent to clients #TEXTO_ENVIO
               # copy method creates a clone of the information so it can be used in other loops

               # Registration confirmation
               janela_texto['titulo_texto'].update('Text registered successfully!',font=('Roboto', 25, 'bold'),text_color='#00FF00')

               # First message
               texto_pronto=valores_texto['texto']
               texto_cadastrado=copy.deepcopy(texto_pronto)

               # Second message
               texto_pronto1=valores_texto['texto1']
               texto_cadastrado1=copy.deepcopy(texto_pronto1)

               # Third message
               texto_pronto2=valores_texto['texto2']
               texto_cadastrado2=copy.deepcopy(texto_pronto2)

               # Clear text field
               print(texto_cadastrado)
               janela_texto['texto'].update('')
               janela_texto['texto1'].update('')
               janela_texto['texto2'].update('')

            if evento_texto == 'cancelar_texto':
              texto_cadastrado=None
              texto_cadastrado1=None
              texto_cadastrado2=None
              janela_texto.close()
              time.sleep(0.3)
              janela_principal.un_hide()

       # Wait for contacts from imobzi
       if evento_principal =='Aguardar Contatos':
           janela_principal.hide()

           layout_aviso_aguardar = [
               [sg.Text('Don\'t forget to check if\nyour accounts are logged in!', text_color='#FFFFFF',
                        justification='center', font=('Century Gothic', 20, 'bold'), background_color='#000000')],
               [sg.Column([[sg.Button('ok', key='ok_aguardar', button_color=('#000000', '#FFFFFF'), font=fonte_geral_botões, size=(10, 1))],
                           [sg.Button('Back', key='voltar_aguardar', button_color=('#000000', '#FFFFFF'), font=fonte_geral_botões, size=(10, 1))]],
                          justification='center')]
           ]

           janela_aviso_aguardar = sg.Window('Wait for Contacts', layout_aviso_aguardar, size=(470, 180))

           while True:
               evento_aguardar, valores_aguardar = janela_aviso_aguardar.read()

               if evento_aguardar == sg.WIN_CLOSED:
                   janela_aviso_aguardar.close()
                   janela_principal.un_hide()
                   break

               if evento_aguardar == 'ok_aguardar':
                  janela_aviso_aguardar.close()
                  time.sleep(0.3)

                  driver = iniciar_driver()
                  driver.get('https://my.imobzi.com/index.html#/home')
                  time.sleep(3)

                  while True:
                      try:
                          while True:
                              # Remember to increase this time when running the program
                              time.sleep(3)

                              # Refresh page (imobzi logo)
                              espere_e_clique(driver, By.XPATH, '//div[@class="imobzi-mascot"]')

                              for c in range(1):
                                  print('entered the loop')

                                  # Click money icon
                                  espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-menu/div/ion-content/div[2]/ul/li[6]/div/div/button")

                                  # Open menu
                                  espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-nav/page-deals/ion-header/ion-navbar/div[2]/div")
                                  print('found the menu')
                                  time.sleep(0.5)

                                  # Select rental sector
                                  espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ion-popover/div/div[2]/div/menu-popover/ion-list/div[1]/div")
                                  time.sleep(3)

                                  # Click on contact (This is the waiting loop - if no contact has arrived it will stay in this loop alternating between pattern 1 and 2 and page refresh)
                                  time.sleep(5)
                                  visão_contato = None
                                  cont = 0
                                  while visão_contato is None:
                                      try:
                                          # Can't be a function otherwise it won't go to the except
                                          visão_contato = driver.find_element(By.XPATH,'//div[@id="6586429703454720" and contains(@class, "col-deal ")]//ion-item')

                                      except NoSuchElementException:
                                          cont += 1
                                          if cont % 2 == 0:
                                              time.sleep(5)
                                              driver.refresh()
                                              time.sleep(30)
                                              print('pattern 1')
                                          else:
                                              time.sleep(30)
                                              # Refresh page (imobzi logo)
                                              espere_e_clique(driver, By.XPATH, '//div[@class="imobzi-mascot"]')
                                              time.sleep(10)
                                              # Click money icon
                                              espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-menu/div/ion-content/div[2]/ul/li[6]/div/div/button")
                                              print('pattern 2')

                                  # Click on contact
                                  espere_e_clique(driver, By.XPATH, '//div[@class="deal-item-container deal-stagnant"]')
                                  print('clicked on contact')

                                  # Click WhatsApp
                                  espere_e_clique(driver, By.XPATH, '//button[@class="whatsapp-button"]')

                                  # Click 'send message'
                                  espere_e_clique(driver, By.XPATH, '/html/body/ion-app/ion-modal/div/whatsapp-sharing/ion-footer/div/div/div[2]/button')

                                  # Get window handles
                                  handles = driver.window_handles
                                  pag = []
                                  # Print identifiers
                                  for handle in handles:
                                      c = handle
                                      pag.append(handle)
                                      print(handle)

                                  # Switch focus to WhatsApp window
                                  driver.switch_to.window(handle)

                                  # For incorrect link
                                  link_incorreto = None
                                  try:
                                      driver.find_element(By.XPATH, "//h2[contains(@class, '_9vd5') and contains(@class, '_9scb') and contains(text(), 'O link está incorreto. Feche essa janela e tente usar outro link.')]")
                                      time.sleep(2)
                                      driver.close()
                                      time.sleep(0.5)
                                      driver.switch_to.window(pag[0])
                                      time.sleep(2)
                                      # Flag contact as lost
                                      espere_e_clique(driver, By.XPATH, '//button[@id="loss"]')
                                      time.sleep(0.5)
                                      # Open reason tabs
                                      espere_e_clique(driver, By.XPATH, "//div[@class='lost-reason-field']")
                                      time.sleep(0.5)
                                      # Invalid contact data (option)
                                      espere_e_clique(driver, By.XPATH, "//div[@class='popover-list-item']//div[@class='popover-item ' and span='Dados de contato inválido']")
                                      time.sleep(0.5)
                                      # OK arrow
                                      espere_e_clique(driver, By.XPATH, '//button[@form="deal-lost-reason-form" and @type="submit"]')
                                      time.sleep(5)
                                      print( 'Invalid link present, contact flagged and registered as "Invalid Contact Data"')
                                      break

                                  except NoSuchElementException:
                                      pass

                                  # Message button
                                  espere_e_clique(driver, By.XPATH, '//*[@id="action-button"]')
                                  print('got here')

                                  # Use WhatsApp Web
                                  espere_e_clique(driver, By.XPATH, '//*[@id="fallback_block"]/div/div/h4[2]/a/span')
                                  time.sleep(5)

                                  # URL Error
                                  wait = WebDriverWait(driver,30)
                                  alert_xpath = '//button[@data-testid="popup-controls-ok"]'  # Replace with your alert XPath
                                  try:
                                      alert = wait.until(EC.presence_of_element_located((By.XPATH, alert_xpath)))

                                      if alert:
                                          print('Found contact with invalid URL')
                                          driver.close()
                                          time.sleep(0.5)
                                          driver.switch_to.window(pag[0])
                                          time.sleep(2)

                                          # Flag contact as lost
                                          espere_e_clique(driver, By.XPATH, '//button[@id="loss"]')
                                          time.sleep(0.5)
                                          # Open reason tabs
                                          espere_e_clique(driver, By.XPATH, "//div[@class='lost-reason-field']")
                                          time.sleep(0.5)
                                          # Invalid contact data (option)
                                          espere_e_clique(driver, By.XPATH, "//div[@class='popover-list-item']//div[@class='popover-item ' and span='Dados de contato inválido']")
                                          time.sleep(0.5)
                                          # OK arrow
                                          espere_e_clique(driver, By.XPATH,'//button[@form="deal-lost-reason-form" and @type="submit"]')
                                          time.sleep(5)
                                          print('contact successfully flagged')
                                          break
                                  except:
                                      pass

                                  # Locate message box (Had to use this structure with while loop because function with WebDriverWait didn't work)
                                  caixa_de_mensagem = None
                                  while caixa_de_mensagem is None:
                                      try:
                                          caixa_de_mensagem = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]')

                                      except NoSuchElementException:
                                          time.sleep(3)

                                  # Click on message box
                                  caixa_de_mensagem.click()
                                  time.sleep(1)

                                  # If no custom message was registered, send default message
                                  if texto_cadastrado == 0:
                                      # Write message
                                      caixa_de_mensagem.send_keys('Hello, My name is Pedro I\'m a real estate agent at Jaguar')
                                      time.sleep(1)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(1)

                                      caixa_de_mensagem.send_keys( 'You contacted us about properties for rent:  https://www.jaguarimoveisguarulhos.com.br/buscar?order=undefined&direction=undefined&availability=rent')
                                      time.sleep(6)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(1)

                                      caixa_de_mensagem.send_keys( 'Do you remember which property you saw? If not, tell me roughly what you need and I\'ll look for you')
                                      time.sleep(1)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(4)
                                  else:
                                      # Write message
                                      caixa_de_mensagem.send_keys(texto_cadastrado)
                                      time.sleep(1)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(1)

                                      caixa_de_mensagem.send_keys(texto_cadastrado1)
                                      time.sleep(6)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(1)

                                      caixa_de_mensagem.send_keys(texto_cadastrado2)
                                      time.sleep(1)
                                      # Send
                                      caixa_de_mensagem.send_keys(Keys.ENTER)
                                      time.sleep(4)

                                  # Close tab
                                  driver.close()
                                  print('no news boss')
                                  time.sleep(3)

                                  # Return to CRM window (handles are stored in a list and referenced the first handle(CRM) with the command pag[0])
                                  driver.switch_to.window(pag[0])

                                  # Update contact to INTEREST area
                                  espere_e_clique(driver, By.XPATH, '//a[contains(text(),"INTERESSE")]')
                                  time.sleep(5)

                                  # Back to contacts
                                  driver.back()

                      except Exception as erro:
                            time.sleep(5)
                            driver.refresh()
                            time.sleep(60)
                            print('error handling pattern inside block')
                            print(erro)
                            print('Refreshed page')

               if evento_aguardar == 'voltar_aguardar':
                    janela_aviso_aguardar.close()
                    time.sleep(0.3)
                    janela_principal.un_hide()

       # Mass message sending
       if evento_principal == 'Enviar Mensagens':
           janela_principal.hide()

           layout_disparo_principal = [
               [sg.Text('Hello! Choose an option:', key='comunicado', background_color='#000000',
                        font=('Lato', 16, 'bold'))],
               [sg.Button('Check', key='conferir_disparo', size=(12, 1), button_color=(cor_botao),
                          font=(fonte_geral_botões))],
               [sg.Button('Send', size=(12, 1), button_color=(cor_botao), font=(fonte_geral_botões)),
                sg.Button('Back', size=(12, 1), button_color=(cor_botao), font=(fonte_geral_botões))],
               [sg.Multiline(key='mensagem', size=(40, 10))],
               [sg.Text(f' Numbers registered: {len(contatos)}', key='cont',
                        font=('Roboto', 10, 'bold'), background_color='#000000')],
               [sg.Button('ok', key='formatar_numero', size=(12, 1), button_color=(cor_botao),
                          font=(fonte_geral_botões))],
               [sg.Text('', key='retorno', background_color='#000000', font=('lato', 12))],
           ]

           janela_disparo = sg.Window('WhatsApp Advertising', layout_disparo_principal, finalize=True)

           # Main message sending loop
           while True:
               evento_disparo, valores_disparo = janela_disparo.read()

               if evento_disparo == sg.WIN_CLOSED:
                janela_disparo.close()
                janela_principal.un_hide()
                break

               # Change "Send" button color when list has a registered number
               if len(contatos) >= 1:
                   janela_disparo['enviar'].update(button_color=('black', '#00FF91'))

               if evento_disparo == 'Conferir Numeros':
                   janela_disparo['retorno'].update(f'Numbers already registered: {contatos}')

               # If pressed send and no message was registered
               if evento_disparo == 'enviar' and len(texto_cadastrado) is None:
                   janela_disparo['retorno']. update('You need to register the message to be sent in:')

               # Back to main window
               if evento_disparo == 'voltar':
                   janela_disparo.close()
                   time.sleep(0.3)
                   janela_principal.un_hide()

               # If pressed send and len of contacts equals zero
               if evento_disparo == 'enviar' and len(contatos) == 0:
                janela_disparo['retorno'].update('No numbers registered')

               # Pressed send and everything is correct
               if evento_disparo == 'enviar' and len(contatos) >= 1:
                   janela_disparo.close()

                   layout_disparo2 = [
                       [sg.Text(f'Messages will be sent', key='aviso1', font=('Lato', 12, 'bold'),
                                background_color='#000000')],
                       [sg.Text(f'Press cancel to stop sending', font=('Lato', 12, 'bold'),
                                background_color='#000000')],
                       [sg.Button('Cancel', key='cancelar', button_color=(cor_botao), size=(12, 1))],
                   ]

                   # Open waiting window
                   janela_disparo2 = sg.Window('Sending messages', layout_disparo2, background_color=cor_de_fundo)

                   # Send messages
                   while len(contatos) >= 1:
                       evento_disparo2, valores_disparo2 = janela_disparo2.read()
                       del contatos[0]
                       time.sleep(30)
                       keyboard.press_and_release('ctrl+w')
                       if len(contatos) <= 0:
                           janela_disparo2['aviso1'].update('All messages have been sent', text_color='green')

                   if evento_disparo2 == 'cancelar':
                     janela_principal.un_hide()
                     break

               if evento_disparo == 'formatar_numero' and len(valores_disparo['mensagem']) > 0:
                   limpo = final(valores_disparo['mensagem'])
                   contatos.extend(limpo)
                   janela_disparo['cont'].update(f'Numbers already registered: {len(contatos)}')
                   janela_disparo['mensagem'].update('')
                   janela_disparo['enviar'].update(button_color=('black', '#00FF91'))

               if evento_disparo == 'conferir_disparo':
                   if len(contatos) == 0:
                       janela_disparo['retorno'].update('No numbers registered')
                   else:
                       janela_disparo['retorno'].update(contatos)

           janela_disparo.close()

       # Return to old clients
       if evento_principal == 'Retornar Aos Clientes':
           janela_principal.hide()

           # Follow-up interface
           layout_retorno = [
               [sg.Text('How many clients do you want to return to?', font=('Lato', 16, 'bold'), background_color='#000000')],
               [sg.Radio('5 clients', 'grupo', key='5', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Radio('10 clients', 'grupo', key='10', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Radio('15 clients', 'grupo', key='15', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Button('Next', font=fonte_geral_botões, button_color=cor_botao, key='enviar_retorno'),
                sg.Button('Back', font=fonte_geral_botões, button_color=cor_botao, key='voltar_retorno')],
           ]

           janela_retorno = sg.Window('Return to clients $', layout_retorno)

           total=0
           while True:
            evento_retorno, valores_retorno = janela_retorno.read()

            if valores_retorno['5']:
                total = 5

            if valores_retorno['10']:
                total = 10

            if valores_retorno['15']:
                total = 15

            # If window is closed
            if evento_retorno == sg.WIN_CLOSED:
              janela_retorno.close()
              janela_principal.un_hide()
              break

            # Back to main window
            if evento_retorno == 'voltar_retorno':
             janela_retorno.close()
             janela_principal.un_hide()
             break

            if evento_retorno == 'enviar_retorno':
                janela_retorno.hide()
                time.sleep(0.5)

                # Next window
                layout_retorno2 = [
                    [sg.Text(' From which client should I start?', font=('Lato', 16, 'bold'),
                             background_color='#000000')],
                    [sg.Radio('From 5th contact', 'grupo_retorno2', key='5°', font=(fonte_geral_botões),
                              background_color='#000000')],
                    [sg.Radio('From 10th contact', 'grupo_retorno2', key='10°', font=(fonte_geral_botões),
                              background_color='#000000')],
                    [sg.Radio('From 15th contact', 'grupo_retorno2', key='15°', font=(fonte_geral_botões),
                              background_color='#000000')],
                    [sg.Radio('From 20th contact', 'grupo_retorno2', key='20°', font=(fonte_geral_botões),
                              background_color='#000000')],
                    [sg.Button('Return $', font=fonte_geral_botões, button_color=cor_botao, key='ok_retorno2'),
                     sg.Button('Back', font=fonte_geral_botões, button_color=cor_botao, key='voltar_retorno2')],
                ]

                janela_retorno2 = sg.Window('Choose first contact', layout_retorno2)

            while True:
                    evento_retorno2, valores_retorno2 = janela_retorno2.read()

                    if evento_retorno2 == 'voltar_retorno2':
                        janela_retorno2.close()
                        janela_retorno.un_hide()

                    if evento_retorno2 == sg.WIN_CLOSED:
                        janela_retorno2.close()
                        janela_retorno.un_hide()
                        break

                    if evento_retorno2 == 'ok_retorno2':
                        janela_retorno2.hide()
                        # End of last window, start execution

                        driver = iniciar_driver()
                        time.sleep(1)
                        driver.get('https://my.imobzi.com/index.html#/initial-page')

                        time.sleep(2)

                        # Click money icon
                        espere_e_clique(driver, By.XPATH,"/html/body/ion-app/ng-component/ion-split-pane/ion-menu/div/ion-content/div[2]/ul/li[6]/div/div/button")

                        # Open menu
                        espere_e_clique(driver, By.XPATH,"/html/body/ion-app/ng-component/ion-split-pane/ion-nav/page-deals/ion-header/ion-navbar/div[2]/div")
                        print('found menu')

                        # Select rental
                        espere_e_clique(driver, By.XPATH,"/html/body/ion-app/ion-popover/div/div[2]/div/menu-popover/ion-list/div[1]/div")
                        time.sleep(3)

                        for c in range(total):
                            time.sleep(5)

                            # Select which contact to start from
                            if valores_retorno2['5°']:
                                # Scroll to desired contact
                                contato = aguardar_carregar(driver, By.XPATH, "//ion-item[starts-with(@class, 'deal-item')]")
                                time.sleep(5)
                                elemento = driver.find_element(By.XPATH, "//div[starts-with(@class,'deal-item-container')][5]")
                                driver.execute_script("arguments[0].scrollIntoView();", elemento)

                                # Click contact
                                time.sleep(3)
                                espere_e_clique(driver, By.XPATH, "//div[starts-with(@class,'deal-item-container')][5]")

                            if valores_retorno2['10°']:
                                # Scroll to desired contact
                                contato = aguardar_carregar(driver, By.XPATH, "//ion-item[starts-with(@class, 'deal-item')]")
                                time.sleep(5)
                                elemento = driver.find_element(By.XPATH, "//div[starts-with(@class,'deal-item-container')][10]")
                                driver.execute_script("arguments[0].scrollIntoView();", elemento)

                                # Click contact
                                time.sleep(3)
                                espere_e_clique(driver, By.XPATH, "//div[starts-with(@class,'deal-item-container')][10]")

                            if valores_retorno2['15°']:
                                # Scroll to desired contact
                                contato = aguardar_carregar(driver, By.XPATH, "//ion-item[starts-with(@class, 'deal-item')]")
                                time.sleep(5)
                                elemento = driver.find_element(By.XPATH, "//div[starts-with(@class,'deal-item-container')][15]")
                                driver.execute_script("arguments[0].scrollIntoView();", elemento)

                                # Click contact
                                time.sleep(3)
                                espere_e_clique(driver, By.XPATH, "//div[starts-with(@class,'deal-item-container')][15]")

                            if valores_retorno2['20°']:
                                # Scroll to desired contact
                                contato = aguardar_carregar(driver, By.XPATH, "//ion-item[starts-with(@class, 'deal-item')]")
                                time.sleep(5)
                                elemento = driver.find_element(By.XPATH, "//div[starts-with(@class,'deal-item-container')][20]")
                                driver.execute_script("arguments[0].scrollIntoView();", elemento)

                                # Click contact
                                time.sleep(3)
                                espere_e_clique(driver, By.XPATH, "//div[starts-with(@class,'deal-item-container')][20]")

                            # Check if contact was already returned to
                            time.sleep(3)
                            try:
                                nome = driver.find_element(By.XPATH,'/html/body/ion-app/ng-component/ion-split-pane/ion-nav/deal-details/div/div/div[2]/div[1]/div[2]/div[2]/div/div/deal-details-data/ion-row[1]/div/deal-contact/ion-item/div[1]/div/ion-label/div/h2')
                                contato = nome.text

                                if contato.lower() == 'ret':
                                    print('Found contact that already received return')
                                    time.sleep(3)
                                    # Update contact
                                    espere_e_clique(driver, By.XPATH, '//a[contains(text(),"FOLLOW-UP")]')
                                    time.sleep(10)

                                    # Click logo
                                    espere_e_clique(driver, By.XPATH, '//div[@class="imobzi-mascot"]')
                                    time.sleep(5)
                                    print('Directed contact to correct area')
                                    continue

                            except NoSuchElementException:
                                pass

                            # Click WhatsApp
                            espere_e_clique(driver, By.XPATH, '//button[@class="whatsapp-button"]')

                            # Click 'send message'
                            espere_e_clique(driver, By.XPATH,'/html/body/ion-app/ion-modal/div/whatsapp-sharing/ion-footer/div/div/div[2]/button')

                            # Get window handles
                            handles = driver.window_handles
                            pag = []
                            # Print identifiers
                            for handle in handles:
                                c = handle
                                pag.append(handle)
                                print(handle)

                            # Switch focus to WhatsApp window
                            driver.switch_to.window(handle)

                            # For incorrect link
                            page_source = driver.page_source
                            if "O link está incorreto. Feche essa janela e tente usar outro link." in page_source:
                                time.sleep(2)
                                driver.close
                                print('Had incorrect link, closed window and continued service')
                                break

                            # Message button
                            espere_e_clique(driver, By.XPATH, '//*[@id="action-button"]')

                            # Use WhatsApp Web
                            espere_e_clique(driver, By.XPATH, '//*[@id="fallback_block"]/div/div/h4[2]/a/span')
                            time.sleep(5)

                            # URL Error
                            wait = WebDriverWait(driver, 30)
                            alert_xpath = '//button[@data-testid="popup-controls-ok"]'  # Replace with your alert XPath
                            try:
                                alert = wait.until(EC.presence_of_element_located((By.XPATH, alert_xpath)))

                                if alert:
                                    print('Found contact with invalid URL')
                                    driver.close()
                                    time.sleep(0.5)
                                    driver.switch_to.window(pag[0])
                                    time.sleep(2)

                                    # Flag contact as lost
                                    espere_e_clique(driver, By.XPATH, '//button[@id="loss"]')
                                    time.sleep(0.5)
                                    # Open reason tabs
                                    espere_e_clique(driver, By.XPATH, "//div[@class='lost-reason-field']")
                                    time.sleep(0.5)
                                    # Invalid contact data (option)
                                    espere_e_clique(driver, By.XPATH,
                                                    "//div[@class='popover-list-item']//div[@class='popover-item ' and span='Dados de contato inválido']")
                                    time.sleep(0.5)
                                    # OK arrow
                                    espere_e_clique(driver, By.XPATH,
                                                    '//button[@form="deal-lost-reason-form" and @type="submit"]')
                                    time.sleep(5)
                                    print('contact successfully flagged')
                                    break
                            except:
                                pass

                            # Locate message box
                            caixa_de_mensagem = None
                            while caixa_de_mensagem is None:
                                try:
                                    caixa_de_mensagem = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]')

                                except NoSuchElementException:
                                    time.sleep(3)

                            # Click on message box
                            caixa_de_mensagem.click()
                            time.sleep(1)

                            # Which message to send (default or custom)
                            if texto_cadastrado is None:
                                # Write message
                                caixa_de_mensagem.send_keys('Hello, I\'m from Jaguar Real Estate')
                                time.sleep(1)
                                # Send message
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(2)

                                # Write message
                                caixa_de_mensagem.send_keys('A while ago we talked about properties for rent, have you found your property yet?')
                                time.sleep(1)
                                # Send message
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(3)

                                # Write message
                                caixa_de_mensagem.send_keys('If you haven\'t found it yet, type "OK" and we\'ll show you other properties')
                                time.sleep(1)
                                # Send message
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(4)
                            else:
                                # Write message
                                caixa_de_mensagem.send_keys(texto_cadastrado)
                                time.sleep(1)
                                # Send
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(1)

                                caixa_de_mensagem.send_keys(texto_cadastrado1)
                                time.sleep(6)
                                # Send
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(1)

                                caixa_de_mensagem.send_keys(texto_cadastrado2)
                                time.sleep(1)
                                # Send
                                caixa_de_mensagem.send_keys(Keys.ENTER)
                                time.sleep(4)

                            # Close tab
                            driver.close()
                            print('Messages delivered')
                            time.sleep(3)

                            # Return to CRM window
                            driver.switch_to.window(pag[0])

                            # Update contact
                            espere_e_clique(driver, By.XPATH, '//a[contains(text(),"FOLLOW-UP")]')
                            time.sleep(10)

                            # Click logo
                            espere_e_clique(driver, By.XPATH, '//div[@class="imobzi-mascot"]')
                            time.sleep(5)

                            # Click money icon
                            espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-menu/div/ion-content/div[2]/ul/li[6]/div/div/button")
                            print('Next contact')

                        # Completion notification layout
                        layout_aviso_fim_retorno = [
                            [sg.Text('Process completed\nHave a Nice Day! ', text_color='#FFFFFF', justification='left',font=('Century Gothic', 20, 'bold'), background_color='#000000')],
                            [sg.Column([[sg.Button('Back', key='voltar_fim_dispensar', button_color=('#000000', '#FFFFFF'), font=fonte_geral_botões, size=(10, 1))]], justification='center')],
                        ]

                        janela_aviso_fim_retorno = sg.Window('Operation Completed', layout_aviso_fim_retorno, size=(370, 150))

                        # Process completion window
                        while True:
                            evento_fim_retorno, valores_fim_retorno = janela_aviso_fim_retorno.read()

                            if evento_fim_retorno == sg.WINDOW_CLOSED:
                                janela_principal.un_hide()
                                break

                            # Back to main window
                            if evento_fim_retorno == 'voltar_fim_dispensar':
                                janela_aviso_fim_retorno.hide()
                                time.sleep(0.1)
                                janela_retorno.un_hide()

       if evento_principal == 'dispensar':
           janela_principal.hide()

           layout_dispensar = [
               [sg.Text('How many clients do you want to dismiss?', font=('Lato', 16, 'bold'), background_color='#000000')],
               [sg.Radio('10 clients', 'grupo', key='10', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Radio('20 clients', 'grupo', key='20', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Radio('30 clients', 'grupo', key='30', font=(fonte_geral_botões), background_color='#000000')],
               [sg.Button('OK', font=fonte_geral_botões, button_color=cor_botao, key='ok_dispensar'),
                sg.Button('Back', font=fonte_geral_botões, button_color=cor_botao, key='voltar_dispensar')],
           ]

           janela_dispensar = sg.Window('Dismiss returned contacts', layout_dispensar)

           total_dispensar = 0
           while True:
               evento_dispensar, valor_dispensar = janela_dispensar.read()

               if evento_dispensar == sg.WIN_CLOSED:
                   janela_dispensar.close()
                   janela_principal.un_hide()
                   break

               if valor_dispensar['10'] == True:
                   total_dispensar = 10

               if valor_dispensar['20'] == True :
                   total_dispensar = 20

               if valor_dispensar['30'] == True:
                   total_dispensar = 30

               if evento_dispensar == 'ok_dispensar':
                   janela_dispensar.hide()
                   driver = iniciar_driver()
                   time.sleep(0.5)
                   driver.get('https://my.imobzi.com/index.html#/home')
                   time.sleep(3)

                   while True:
                       for c in range(total_dispensar):
                           time.sleep(5)
                           # Return to start
                           espere_e_clique(driver, By.XPATH, '//div[@class="imobzi-mascot"]')

                           # Click money icon
                           espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-menu/div/ion-content/div[2]/ul/li[6]/div/div/button")

                           # Open menu
                           espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ng-component/ion-split-pane/ion-nav/page-deals/ion-header/ion-navbar/div[2]/div")
                           print('found menu')
                           time.sleep(0.5)

                           # Select rental sector
                           espere_e_clique(driver, By.XPATH, "/html/body/ion-app/ion-popover/div/div[2]/div/menu-popover/ion-list/div[1]/div")
                           time.sleep(10)

                           # CLICK ON CONTACT
                           espere_e_clique(driver, By.XPATH, '//*[@id="6542798388985856"]')
                           time.sleep(1)

                           # Flag contact as lost
                           espere_e_clique(driver, By.XPATH, '//button[@id="loss"]')
                           time.sleep(1)

                           # Open reason tabs
                           espere_e_clique(driver, By.XPATH, "//div[@class='lost-reason-field']")
                           time.sleep(1)

                           # Postponed purchase
                           espere_e_clique(driver, By.XPATH,"/html/body/ion-app/ion-popover/div/div[2]/div/menu-popover/ion-list/div[1]/div")

                           # OK arrow
                           espere_e_clique(driver, By.XPATH, '//button[@form="deal-lost-reason-form" and @type="submit"]')
                           time.sleep(8)
                           print('contact successfully flagged')

                       layout_aviso_fim_dispensar = [
                           [sg.Text('Process completed\nHave a Nice Day! ', text_color='#FFFFFF', justification='left', font=('Century Gothic', 20, 'bold'), background_color='#000000')],
                           [sg.Column([[sg.Button('Back', key='voltar_fim_dispensar', button_color=('#000000', '#FFFFFF'), font=fonte_geral_botões, size=(10, 1))]], justification='center')]
                       ]

                       janela_aviso_fim_dispensar = sg.Window('Operation Completed', layout_aviso_fim_dispensar, size=(370, 150))

                       while True:
                           evento_fim_dispensar, valores_fim_dispensar = janela_aviso_fim_dispensar.read()

                           if evento_fim_dispensar == sg.WINDOW_CLOSED:
                               janela_principal.un_hide()
                               break

                           if evento_fim_dispensar == 'voltar_fim_dispensar':
                               janela_aviso_fim_dispensar.close()
                               time.sleep(0.2)
                               janela_dispensar.un_hide()

               if evento_dispensar == 'voltar_dispensar':
                   janela_dispensar.close()
                   time.sleep(0.2)
                   janela_principal.un_hide()
                   break

       if evento_principal == 'login':
               janela_principal.hide()
               time.sleep(1)
               driver = iniciar_driver()

               # Instructions window layout
               layout_login = [
                   [sg.Text(
                       'Minimize this window and manually login to WhatsApp\nand Imobzi in the minimized window, when finished click "Ok".',
                       justification='center', font=('Lato', 16, 'bold'), background_color='#000000')],
                   [sg.Column([[sg.Button('Ok', key='ok', button_color=cor_botao, font=fonte_geral_botões, size=(10, 1))]],
                              justification='center')]
               ]

               # Start Chrome
               driver.get('https://web.whatsapp.com/')
               time.sleep(5)

               # Minimize window
               driver.minimize_window()
               time.sleep(1)

               # Open instructions window
               janela_login = sg.Window('Login to your account', layout_login, size=(670, 125))

               while True:
                evento_login, valores_login = janela_login.read()

                if evento_login == sg.WINDOW_CLOSED:
                  janela_login.close()
                  break

                if evento_login == 'ok':
                    janela_login.close()
                    driver.close()
                    time.sleep(0.1)
                    janela_principal.un_hide()

  except NoSuchElementException:
    time.sleep(2)
    driver.refresh()
    time.sleep(5)
    print(erro)
    print('pattern 3')
    print('Refreshed page')

  except requests.exceptions.Timeout:
     time.sleep(120)
     driver.refresh()
     print('Connection timeout. Waiting 120 seconds ')

  except requests.exceptions.ConnectionError:
     time.sleep(120)
     driver.refresh()
     print("Connection error. Waiting 120 seconds.")

  except Exception as erro:
     time.sleep(2)
     driver.refresh()
     time.sleep(5)
     print('pattern 4')
     print(f'Other type of error{erro}')

# End sleep mode block
ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# Add error handling for internet and case of login on another machine

janela_principal.close()


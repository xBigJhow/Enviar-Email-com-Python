from asyncio import events
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Literal
import PySimpleGUI as sg
import smtplib


class CreateWindow():
    def __init__(self):

        # Trocando o tema da aplicação
        sg.theme(sg.user_settings_get_entry('theme', None))
        
        layout = [
            # Inserção de E-mail e Senha...
            [sg.Text(text='Email: ', size=(10)), sg.Input(default_text='',font='Arial 10', size=(50), k='email_from')],
            [sg.Text(text='Senha: ', size=(10)), sg.Input(password_char='*',default_text='',size=(50), k='pass')],

            # O e-mail enviado será com ou sem assinatura
            [sg.Text(text='Assinatura: ', size=(10)), sg.Radio('Sim', 'signature'), sg.Radio('Não', 'signature')],
            [sg.Text(''*30)],

            # Selecionar o servidor do transmissor do e-mail, por exemplo se o usuário for "@hotmail.com", selecione a opção "Hotmail"
            [sg.Text(text='Selecionar somente o servidor do usuário que vai enviar o e-mail:')],
            [sg.Checkbox(text='Gmail', auto_size_text=True, k='gmail'), sg.Checkbox(text='Hotmail', auto_size_text=True, k='hotmail'), sg.Checkbox('Yahoo', auto_size_text=True, k='yahoo'),sg.Checkbox('Uol',auto_size_text=True, k='uol')],
            [sg.Text(''*30)],

            # Email do destinatário do e-mail.
            [sg.Text(text='Para: ', size=(10)), sg.Input(default_text='', size=(50), k='email_to')],

            # Assunto do E-mail (Header)
            [sg.Text(text='Assunto: ', size=(10)), sg.Input(default_text='', size=(50,40), k='message')],

            # Mensagem do E-mail / Corpo do E-mail
            [sg.Text(text='Mensagem: ', size=(10)),sg.Multiline('', size=(48,5), k='describe')],
            [sg.Text(''*30)],

            # Aqui é somente um output que mostrará se foi possível ou não o envio do e-mail.
            [sg.Text('                      Não escreva nada na linha abaixo...')],
            [sg.Text(text='Verificação: ', size=(10)), sg.Output(size=(48,2), background_color='Grey', text_color='white')],
            [sg.Text(''*30)],
            [sg.Button(button_text='Enviar', k='send'), sg.Button(button_text='Cancelar', k='cancel'), sg.Button(button_text='Estilos', k='theme_color')],
        ]
        # Gerando a tela.
        self.py_windows = sg.Window(title='Send Mail', layout=layout, finalize=True, resizable=True)
        

    def StartGUI(self):
        while True:
            self.events, self.values = self.py_windows.Read()

            # Caso botão "X" de fechar a janela ou botão de cancelar sejam pressionados, para o programa.
            if self.events in (sg.WINDOW_CLOSED, 'cancel'):
                break

            # Caso o botão  "Enviar seja pressionado... fará o processo de enviar o e-mail"
            elif self.events == 'send':
                try:
                    msg = MIMEMultipart() 
                    password = self.values['pass']
                    msg['from'] = self.values['email_from']
                    msg['to'] = self.values['email_to']
                    msg['Subject'] = self.values['message']

                    # Verifica se a assinatura foi pressionado sim ou não...
                    if self.values[0] == True:
                        # Nesta variavel insira a assinatura desejada
                        msg['describe'] = self.values['describe'] + """

                        """
                    else:
                        # Caso não queira assinatura, o corpo do e-mail terá somente a mensagem, sem assinatura.
                        msg['describe'] = self.values['describe']
                    
                    
                    msg.attach(MIMEText(_text=(msg['describe'])))
                    gmail = self.values['gmail']
                    hotmail = self.values['hotmail']
                    yahoo = self.values['yahoo']
                    uol = self.values['uol']

                    # Servidores
                    if hotmail:   
                        #print('Servidor Hotmail')
                        server = smtplib.SMTP('smtp-mail.outlook.com: 25')
                    elif gmail:
                        server = smtplib.SMTP('smtp-relay.gmail.com: 587')
                        #print('Servidor Gmail')
                    elif yahoo:
                        server = smtplib.SMTP('smtp.gmail.com: 587')
                       #print('Servidor Yahoo')
                    elif uol:
                        server = smtplib.SMTP('smtps.uol.com.br')
                        #print('Servidor Uol')
                    
                    # Segurança TLS
                    server.starttls()

                    # Logando no servidor com e-mail e senha do remetente
                    server.login(user=msg['from'], password=password)

                    # Enviando Email, com remetente, destinatário e mensagem.
                    server.sendmail(msg['From'], msg['To'],msg.as_string())

                    # Caso deu certo, aparecerá esta mensagem...
                    print("Email enviado com Sucesso")

                    # Desconectando do servidor 
                    server.quit()
                except smtplib.SMTPException:
                    
                    # Caso não der certo, aparecerá está mensagem. 
                    print("ERRO! Falha no envio...")

            # Modificar o Tema da aplicação
            elif self.events in ('theme_color'):
                self.evs, self.vals = sg.Window('Choose Theme', [[sg.Combo(sg.theme_list(), k='-THEME LIST-'), sg.OK(), sg.Cancel(button_text='Cancelar')]]).read(close=True)
                if self.evs == 'OK':
                    self.py_windows.Close()
                    sg.user_settings_set_entry('theme', self.vals['-THEME LIST-'])
                    CreateWindow().StartGUI()  


# Criando e Iniciando a Janela                                  
py_windows = CreateWindow().StartGUI()                       

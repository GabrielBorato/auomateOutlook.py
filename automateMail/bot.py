####################imports###################
# Import for the Desktop Bot
from botcity.core import DesktopBot
# Import for the Web Bot
from botcity.web import WebBot, Browser, By
from webdriver_manager.chrome import ChromeDriverManager
from botcity.plugins.ms365.credentials import MS365CredentialsPlugin, Scopes
from botcity.plugins.ms365.outlook import MS365OutlookPlugin
from botcity.plugins.excel import BotExcelPlugin
bot_excel = BotExcelPlugin()

service = MS365CredentialsPlugin(
    client_id='24efa602-9a99-452d-92cf-45233e3ea674',
    client_secret='2An8Q~QR-sMjyXRvpWL8Af6qe2eTN3PfOXTMIbLe',
)
scopes_list = [Scopes.BASIC, Scopes.FILES_READ_WRITE_ALL, Scopes.MAIL_READ_WRITE, Scopes.MAIL_READ_WRITE, Scopes.MAIL_SEND]
service.authenticate(scopes=scopes_list)
outlook = MS365OutlookPlugin(service_account=service)
#Import for integration with BotCity Maestro SDK
# from botcity.maestro import *
# BotMaestroSDK.RAISE_NOT_CONNECTED = False
####################Trade-Outlook#######################
bot_excel.read('K:\\Contratos e Trade\\T.I\\baseDeDadosConsumoBot.xlsx')
fornecedores = list(bot_excel.get_column(column="B"))
body1 =""
def retorna_sequencial_fornecedor(fornecedores):
    for fornecedor in fornecedores:
        yield fornecedor
for componenteFornecedor in retorna_sequencial_fornecedor(fornecedores):
    body1 += f"{componenteFornecedor}\n"
##print(body1)
#########################################################    
contrapartidas = bot_excel.get_column(column="H")
body2 =""
def retorna_sequencial_contrapartida(contrapartida):
    for contrapartida in contrapartidas:
        yield contrapartida
for componenteContrapartida in retorna_sequencial_contrapartida(contrapartidas):
    body2 += f"{componenteContrapartida}\n"
#########################################################   
lojas = bot_excel.get_column(column="I")
body3 =""
def retorna_sequencial_lojas(lojas):
    for loja in lojas:
        yield loja
for componenteLoja in retorna_sequencial_lojas(lojas):
    body3 += f"{componenteLoja}\n"
#########################################################
mesAcao = bot_excel.get_column(column="J")
body4 =""
def retorna_sequencial_mesAcao(mesAcao):
    for acao in mesAcao:
        yield acao
for componenteAcao in retorna_sequencial_mesAcao(mesAcao):
    body4 += f"{componenteAcao}\n"
#########################################################
contrapartidaqtde = bot_excel.get_column(column="K")
body5 =""
def retorna_sequencial_contrapartidaqtde(contrapartidaqtde):
    for qtdee in contrapartidaqtde:
        yield qtdee
for contrapartidaqtdee in retorna_sequencial_contrapartidaqtde(contrapartidaqtde):
    body5 += f"{contrapartidaqtdee}\n"
#########################################################
email = bot_excel.get_column(column="L")
body6 =""
def retorna_sequencial_email(email):
    for mail in email:
        yield mail
for mailmail in retorna_sequencial_email(email):
    body6 += f"{mailmail}\n"
########################################################
dimensoes = bot_excel.get_column(column="M")
body7 =""
def retorna_sequencial_dimensoes(dimensoes):
    for dim in dimensoes:
        yield dim
for dimensoes in retorna_sequencial_dimensoes(dimensoes):
    body7 += f"{dimensoes}\n"
#########################################################
datas = bot_excel.get_column(column="N")
body8 =""
def retorna_sequencial_datas(datas):
    for date in datas:
        yield date
for datas in retorna_sequencial_datas(datas):
    body8 += f"{datas}\n"
#########################################################
to = [body6]
subject = "Requisição de arte trade Grupo Koch"
files = ['K:\\Contratos e Trade\\T.I\\Execução Plano trade.png']
body = "Opa fiote"

# Enviando a mensagem de e-mail
outlook.send_message(subject, body, to, attachments=files)

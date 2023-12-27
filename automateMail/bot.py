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

def ler_dados_excel():
    return {
        "fornecedores": bot_excel.get_column(column="B")[1:],
        "contrapartidas": bot_excel.get_column(column="H")[1:],
        "lojas": bot_excel.get_column(column="I")[1:],
        "mesAcoes": bot_excel.get_column(column="J")[1:],
        "contrapartidaqtde": bot_excel.get_column(column="K")[1:],
        "dimensoes": bot_excel.get_column(column="M")[1:],
        "datas": bot_excel.get_column(column="N")[1:],
        "emails": bot_excel.get_column(column="L")[1:]
    }

dados = ler_dados_excel()
subject = "Requisição de arte trade Grupo Koch"
files = ['K:\\Contratos e Trade\\T.I\\Execução Plano trade.png']

for fornecedor, data, contrapartida, loja, mesAcao, contrapartidaqtde1, dimensao, email in zip(
        dados["fornecedores"], dados["datas"], dados["contrapartidas"],
        dados["lojas"], dados["mesAcoes"], dados["contrapartidaqtde"],
        dados["dimensoes"], dados["emails"]
):
    body = f"Olá {fornecedor} {data}{contrapartida}{loja}{mesAcao}{contrapartidaqtde1}{dimensao}\n"

    # Enviando a mensagem de e-mail
    outlook.send_message(subject, body, [email], attachments=files)

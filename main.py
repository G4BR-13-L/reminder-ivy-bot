import requests
import time
import json
import os

from xlrd import open_workbook, xldate_as_tuple
from datetime import date, datetime, timedelta
from calendar import monthrange

os.environ['TZ'] = 'America/Sao_Paulo'
time.tzset()

class Event(object):
    def __init__(self, date, course, task, desc):
        self.id = id
        self.date = date
        self.course = course
        self.task = task
        self.desc = desc

    def __lt__(self, other):
         return self.date < other.date

    def __str__(self):
        return("*Data:* {0}\n"
               "*Disciplina:* {1}\n"
               "*Atividade:* {2}\n"
               "*Descricao:* {3}\n\n"
               .format(self.date.strftime("%d/%m/%Y"), self.course, self.task,
                       self.desc))

class TelegramBot:
    def __init__(self):
        token = 'INSERT_TOKEN_HERE'
        self.url_base = f'https://api.telegram.org/bot{token}/'

    def Iniciar(self):
        update_id = None
        while True:
            atualizacao = self.obter_novas_mensagens(update_id)
            dados = atualizacao["result"]
            if dados:
                for dado in dados:
                    update_id = dado['update_id']
                    mensagem = str(dado["message"]["text"])
                    chat_id = dado["message"]["from"]["id"]
                    first = int(
                        dado["message"]["message_id"]) == 1
                    resposta = self.criar_resposta(
                        mensagem, first)
                    self.responder(resposta, chat_id)

    # Obter mensagens
    def obter_novas_mensagens(self, update_id):
        link_requisicao = f'{self.url_base}getUpdates?timeout=100'
        if update_id:
            link_requisicao = f'{link_requisicao}&offset={update_id + 1}'
        resultado = requests.get(link_requisicao)
        return json.loads(resultado.content)

    def last_day_of_month(self, date_value):
        return date_value.replace(day = monthrange(date_value.year, date_value.month)[1])

    def last_day_of_week(self, date_value):
        start = date_value - timedelta(days=date_value.weekday())
        return start + timedelta(days=6)

    # Criar uma resposta
    def criar_resposta(self, mensagem, first):
        resp = ""

        today = datetime.today().strftime('%d/%m/%Y')
        today = datetime.strptime(today, "%d/%m/%Y")

        if mensagem == '/mes':
            wb = open_workbook('Book.xlsx')

            end = self.last_day_of_month(today)

            for sheet in wb.sheets():
                nrows = sheet.nrows
                ncols = sheet.ncols

                itens = []

                for row in range(1, nrows):
                    item = ""

                    Data = (sheet.cell(row,0).value)

                    try:
                      d = xldate_as_tuple(Data, wb.datemode)
                      #convert date tuple in yy-mm-dd format
                      d = datetime(*(d[0:3]))
                      Ex = d.strftime("%d/%m/%Y")
                      Data = Ex
                    except Exception as e:
                      pass
                    else:
                      pass
                    finally:
                      pass

                    Data = str(Data)

                    Disciplina = (sheet.cell(row,1).value)
                    Atividade = (sheet.cell(row,2).value)
                    Descricao = (sheet.cell(row,3).value)

                    DtObj = datetime.strptime(Data, "%d/%m/%Y")
                    
                    EvArr = [DtObj, Disciplina, Atividade, Descricao]

                    if (DtObj >= today and DtObj <= end):
                        itens.append(Event(*EvArr))

                itens.sort()

                for item in itens:
                    resp += item.__str__()

            return resp

        if mensagem == '/semana':
            wb = open_workbook('Book.xlsx')

            end = self.last_day_of_week(today)

            for sheet in wb.sheets():
                nrows = sheet.nrows
                ncols = sheet.ncols

                itens = []

                for row in range(1, nrows):
                    item = ""

                    Data = (sheet.cell(row,0).value)

                    try:
                      d = xldate_as_tuple(Data, wb.datemode)
                      #convert date tuple in yy-mm-dd format
                      d = datetime(*(d[0:3]))
                      Ex = d.strftime("%d/%m/%Y")
                      Data = Ex
                    except Exception as e:
                      pass
                    else:
                      pass
                    finally:
                      pass

                    Data = str(Data)

                    Disciplina = (sheet.cell(row,1).value)
                    Atividade = (sheet.cell(row,2).value)
                    Descricao = (sheet.cell(row,3).value)

                    DtObj = datetime.strptime(Data, "%d/%m/%Y")
                    
                    EvArr = [DtObj, Disciplina, Atividade, Descricao]

                    if (DtObj >= today and DtObj <= end):
                        itens.append(Event(*EvArr))

                itens.sort()

                for item in itens:
                    resp += item.__str__()

            return resp

        if first == True or mensagem in ('/requisitos'):
            return f'''Olá, tenho alguns requisitos e ideias para voce:{os.linesep}{os.linesep}Ideias:{os.linesep}- Pesquisar atividades somente de datas presentes e futuras{os.linesep}- Responder a trabalhos semanais e mensais{os.linesep}- chamar pelo nome'''
        elif mensagem == '/structuretest':
          return f'''Data: 12 de abril;{os.linesep}Disciplina: Interfaces web.{os.linesep}Atividade: Quiz{os.linesep}Descrição: quiz sobre css{os.linesep}'''
        else:
            return f'''Como vai você? :){os.linesep}{os.linesep}Digite /requisitos para entender o que preciso.'''
        

    # Responder
    def responder(self, resposta, chat_id):
        link_requisicao = f'{self.url_base}sendMessage?chat_id={chat_id}&text={resposta}&parse_mode=markdown'
        requests.get(link_requisicao)


bot = TelegramBot()
bot.Iniciar()

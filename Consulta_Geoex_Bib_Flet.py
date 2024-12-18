from requests import post
from gspread import authorize
from pandas import DataFrame
from json import load, dump
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import os

# Abrir credenciais do Google Sheets
#scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
scope = 'https://spreadsheets.google.com/feeds'
creds = ServiceAccountCredentials.from_json_keyfile_name('_internal/jimmy.json', scope)
gs = authorize(creds)
url_geo = 'https://geoex.com.br/api/Cadastro/ConsultarProjeto/'
url_pasta = 'https://geoex.com.br/api/ConsultarProjeto/EnvioPasta/Itens'

def abre(arquivo):
    with open(arquivo) as dados:
        jason = load(dados)
    return jason

#Abre as planilhas registradas
planilhas = abre('_internal/meses.json')
meses = [a for a in planilhas]
mes = meses[-1]
juncao = '1N0G-2y4fFniqwmIn716vakAWdZJRhV2cCy9yHHcINP8'
abas_pastas = ['CAPEX_2024','OPEX_2024']

#Abre o cookie para o GEOEX
data = abre('_internal/cookie.json')
cookie = data['cookie']
gxsessao = data['gxsessao']
useragent = data['useragent']

def hora_atual():
    data_e_hora_atuais = datetime.now()
    data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M')

    return data_e_hora_em_texto

def escreve_json(arquivo, dicionario):
    with open(arquivo, "w") as outfile: 
        dump(dicionario, outfile)

def procura_projeto(projeto, cookie=cookie, gxsessao=gxsessao, useragent=useragent):
    #print(cookie, '\n', gxsessao, '\n', useragent)
    id_projeto = ''
    r = ''
    projeto = str(projeto).strip()
    fim = False

    url = url_geo+'Item'
    header = {
        'Cookie': cookie,
        'User-Agent': useragent,
        'Gxsessao': gxsessao
    }
    body = {
        'id':projeto
    }

    while True:
        try:
            r = post(url = url, headers = header, json = body)
            if r.status_code!=200:
                print("\033[K", projeto+' Erro na requisição: Code: '+str(r.status_code)+', Reason: '+str(r.reason), end='\r')
                fim = True
                continue
            if fim:
                fim = False
                #print('')
            r = r.json()
            break
        except Exception as e:
            print(hora_atual())
            print('Exception', e)
            print('Requisição', r)
            print('Projeto ', projeto)
            raise Exception('Não foi possível acessar a página do GEOEX.')

    if r['Content'] != None:
        id_projeto = r['Content']['ProjetoId']
    elif r['IsUnauthorized']:
        print(r)
        raise Exception('Cookie inválido! Não autorizado')

    return id_projeto

def consulta_medicao_geoex(projeto, idmedicao, cookie=cookie, gxsessao=gxsessao, useragent=useragent):
    validacao = ''
    data_postagem = ''
    data_atesto = ''
    data_validacao = ''
    r = None
    idmedicao = str(idmedicao).strip()
    fim = False

    idprojeto = procura_projeto(projeto, cookie, gxsessao, useragent)

    if idprojeto == '':
        return validacao, data_postagem, data_atesto, data_validacao

    url = 'https://geoex.com.br/api/ConsultarProjeto/Postagem/Itens'
    header = {
        'cookie': cookie,
        'User-Agent': useragent,
        'Gxsessao': gxsessao
    }
    body = {
        'ProjetoId':idprojeto,
        'Rejeitadas':True,
        'Arquivadas':False,
        'Paginacao':{'TotalPorPagina':'200','Pagina':'1'},
        'Search':idmedicao
    }

    while True:
        try:
            r = post(url = url, headers = header, json = body)
            if r.status_code!=200:
                print("\033[K", projeto+' Erro na requisição: Code: '+str(r.status_code)+', Reason: '+str(r.reason), end='\r')
                fim = True
                continue
            if fim:
                fim = False
                #print('')
            r = r.json()
            break
        except Exception as e:
            print(hora_atual())
            print('Exception', e)
            print('Requisição', r)
            print('Projeto ', projeto)
            raise Exception('Não foi possível acessar a página do GEOEX.')
    
    if r['Content'] != None:
        for i in r['Content']['Items']:
            #print('Buscando medição ' + idmedicao + '.')
            #print(i['Serial'], i['Status'], idmedicao, i['Serial']==idmedicao)
            if i['Serial']==idmedicao:
                #print('Medição ' + idmedicao + ' encontrada.')
                if i['Status'] != None:
                    validacao = i['Status']
                if i['Data'] != None:
                    data_postagem = datetime.fromisoformat(i['Data']).date().isoformat()
                if i['DataAtesto'] != None:
                    data_atesto = datetime.fromisoformat(i['DataAtesto']).date().isoformat()
                if i['DataValidacao'] != None:
                    data_validacao = datetime.fromisoformat(i['DataValidacao']).date().isoformat()
                return validacao, data_postagem, data_atesto, data_validacao

        #print('Medição ' + idmedicao + ' não encontrada no projeto ' + projeto +'.')
        validacao = 'Não Encontrado'

    return validacao, data_postagem, data_atesto, data_validacao

def acha_nome(nome):
    inicio = nome.find("'")
    inicio += 1
    fim = nome[inicio:].find("'")
    fim = inicio + fim
    return nome[inicio:fim]

def atualiza_medicao(planilha, sh, mes, progresso, porcentagem, nomeplanilha, intervalo, cookie=cookie, gxsessao=gxsessao, useragent=useragent):
    nomeplanilha.value = f'Lendo as medições de {planilha} - {mes}'
    nomeplanilha.update()
    sheet = sh.worksheet(planilha).get_all_values()
    sheet = DataFrame(sheet, columns = sheet.pop(0))
    tamanho = sheet.shape[0]
    #print('Atualizando status das medições de ' + planilha)
    valores = []
    a=[]
    idgeoex = 'a'
    cont = 0

    for i,j in enumerate(sheet['ID GEOEX / N° OS']):
        #if i>130:
        #    break
        #if i<118:
        #    continue
        progresso.value = i/sheet.shape[0]
        porcentagem.value =text='{:.2f}%'.format(progresso.value*100)
        progresso.update()
        porcentagem.update()
        #if cont>=20:
        #    break

        c = sheet['PROJETO'][i]
        d = sheet['CÓDIGO SERVIÇO'][i]
        f = sheet['STATUS (GEOEX)'][i]

        if f=='PedidoLancado' or c == 'OBRA' or c == 'EQM' or c=='USO_MUTUO':
        #if c == 'OBRA' or c == 'EQM' or c=='USO_MUTUO':
            a = []
            valores.append(a)
            cont = 0
        elif j == '':
            #print('projeto ' + c)
            if c != '':
                a = ['Não Postado','','']
                #valores.append(a)
                cont = 0
            else:
                if d == '':
                    a = ['']
                    cont += 1
                else:
                    a = ['Não Postado','','']
            valores.append(a)
        elif j == idgeoex:
            valores.append(a)
            cont = 0
        else:
            #print('Atualizando medição ' + j + ' iteração ' + str(i))
            #print('|', end='')
            if j[:2]!= 'PM':
                a = []
            else:
                a = consulta_medicao_geoex(c, j, cookie, gxsessao, useragent)
                #print(j,a)
                if a[0]=='':
                    a = []
                
            valores.append(a)
            cont = 0

        try:
            print("\033[K", str(i)+'/'+str(tamanho)+' - '+c, j, a[0], end="\r")
        except:
            print("\033[K", str(i)+'/'+str(tamanho)+' - '+c, j, a, end="\r")

        idgeoex = j
    
    progresso.value = 1
    porcentagem.value =text='{:.2f}%'.format(progresso.value*100)
    nomeplanilha.value = f'Salvando as medições de {planilha} - {mes}'
    progresso.update()
    porcentagem.update()
    nomeplanilha.update()
    sh.worksheet(planilha).update(intervalo,valores)
    #print('\n' + 'Status das medições de ' + planilha + ' atualizados\n')

def atualiza_planilha(link, progresso, porcentagem, nomeplanilha, intervalo):
    global data, cookie, gxsessao, useragent
    data = abre('_internal/cookie.json')
    cookie = data['cookie']
    gxsessao = data['gxsessao']
    useragent = data['useragent']
    # Chamada p/ atualizar planilha
    if link == '':
        sh = gs.open_by_key(planilhas[meses[-1]])
    else:
        if link in planilhas:
            sh = gs.open_by_key(planilhas.get(link))
        else:
            sh = gs.open_by_key(link)

    lista = sh.worksheets()
    mes = sh.title


    print('\n' + hora_atual() + ': ' + 'Atualizando ' + mes)

    for aba in lista:
        #break
        aba = acha_nome(str(aba))
        if aba[:6] == 'IEM/MP':
            print('\n' + hora_atual() + ': ' + 'Atualizando ' + aba)
            atualiza_medicao(aba, sh, mes, progresso, porcentagem, nomeplanilha, intervalo, cookie, gxsessao, useragent)
            
    nomeplanilha.value = f'Status das medições de {mes} atualizados\n'
    nomeplanilha.update()
            
    print('\n' + hora_atual() + ': ' + mes + ' atualizado!')

def consulta_projeto(projeto, cookie=cookie, gxsessao=gxsessao, useragent=useragent):
    #print(cookie, '\n', gxsessao, '\n', useragent)
    id_projeto, titulo, statusprj, statususuario, statushektor  = '','','','',''
    r = ''
    projeto = str(projeto).strip()
    fim = False

    url = url_geo+'Item'
    header = {
        'Cookie': cookie,
        'User-Agent': useragent,
        'Gxsessao': gxsessao
    }
    body = {
        'id':projeto
    }

    while True:
        try:
            r = post(url = url, headers = header, json = body)
            if r.status_code!=200:
                print("\033[K", projeto+' Erro na requisição: Code: '+str(r.status_code)+', Reason: '+str(r.reason), end='\r')
                fim = True
                continue
            if fim:
                fim = False
                #print('')
            r = r.json()
            break
        except Exception as e:
            print(hora_atual())
            print('Exception', e)
            print('Requisição', r)
            print('Projeto ', projeto)
            raise Exception('Não foi possível acessar a página do GEOEX.')

    if r['Content'] != None:
        id_projeto = r['Content']['ProjetoId']
        if r['Content']['Titulo'] != None:
            titulo = r['Content']['Titulo']
        try:
            if r['Content']['StatusProjeto']['Descricao'] != None:
                statusprj = r['Content']['StatusProjeto']['Descricao']
            if r['Content']['StatusUsuario']['Descricao'] != None:
                statususuario = r['Content']['StatusUsuario']['Descricao']
            if r['Content']['GseProjeto']['Status']['Nome'] != None:
                statushektor = r['Content']['GseProjeto']['Status']['Nome']
        except:
            pass
    elif r['IsUnauthorized']:
        print(r)
        raise Exception('Cookie inválido! Não autorizado')

    return id_projeto, titulo, statusprj, statususuario, statushektor

def consulta_pasta(idprojeto, cookie=cookie, gxsessao=gxsessao, useragent=useragent):
    #print(cookie, '\n', gxsessao, '\n', useragent)
    statusceite, obsaceite, serial= 'NÂO POSTADO','',''
    r = ''
    fim = False

    url = url_pasta
    header = {
        'Cookie': cookie,
        'User-Agent': useragent,
        'Gxsessao': gxsessao
    }
    body = {
        'ProjetoId':idprojeto
    }

    while True:
        try:
            r = post(url = url, headers = header, json = body)
            if r.status_code!=200:
                print("\033[K", str(idprojeto)+' Erro na requisição: Code: '+str(r.status_code)+', Reason: '+str(r.reason), end='\r')
                fim = True
                continue
            if fim:
                fim = False
                #print('')
            r = r.json()
            break
        except Exception as e:
            print(hora_atual())
            print('Exception', e)
            print('Requisição', r)
            print('IDProjeto ', idprojeto)
            raise Exception('Não foi possível acessar a página do GEOEX.')

    try:
        if r['Content'] != None:
            if len(r['Content']['Items'])>0:
                if r['Content']['Items'][0]['HistoricoStatusId'] != None:
                    if r['Content']['Items'][0]['HistoricoStatusId']==30:
                        statusceite="ACEITO"
                    elif r['Content']['Items'][0]['HistoricoStatusId']==31:
                        statusceite="ACEITO COM RESTRIÇÕES"
                    elif r['Content']['Items'][0]['HistoricoStatusId']==32:
                        statusceite="REJEITADO"
                    elif r['Content']['Items'][0]['HistoricoStatusId']==1:
                        statusceite="CRIADO"
                    elif r['Content']['Items'][0]['HistoricoStatusId']==6:
                        statusceite="CANCELADO"
                    elif r['Content']['Items'][0]['HistoricoStatusId']==35:
                        statusceite="VALIDADO"
                    else:
                        statusceite=str(r['Content']['Items'][0]['HistoricoStatusId'])
                if r['Content']['Items'][0]['Observacao'] != None:
                    obsaceite = r['Content']['Items'][0]['Observacao']
                if r['Content']['Items'][0]['Serial'] != None:
                    serial = r['Content']['Items'][0]['Serial']
        elif r['IsUnauthorized']:
            print(r)
            raise Exception('Cookie inválido! Não autorizado')
    except Exception as e:
        print(e)
        print(r)
        raise Exception('Falha na requisição.')

    return statusceite, obsaceite, serial

def atualiza_pasta(infopastas, progressopasta, porcentagempasta, data):
    sh = gs.open_by_key(juncao)
    print('\n' + hora_atual() + ': ' + 'Atualizando Pastas')
    
    global cookie, gxsessao, useragent
    
    cookie = data['cookie']
    gxsessao = data['gxsessao']
    useragent = data['useragent']

    valores = [[],[]]
    a,b=[],[]

    lista = sh.worksheets()

    for aba in lista:
        if aba.title in abas_pastas:
            valores = [[],[]]
            a,b=[],[]
            print(hora_atual() + ': ' + 'Atualizando ' + aba.title)
            infopastas.value = f'Atualizando status das pastas {aba.title}\n'
            infopastas.update()

            sheet = sh.worksheet(aba.title).get_all_values()
            sheet = DataFrame(sheet, columns = sheet.pop(0))
            sheet = sheet[sheet['PROJETO']!=""]
            tamanho = sheet.shape[0]

            for i,j in enumerate(sheet['PROJETO']):
                progressopasta.value = i/tamanho
                porcentagempasta.value =text='{:.2f}%'.format(progressopasta.value*100)
                progressopasta.update()
                porcentagempasta.update()
                
                #print("\033[K", str(i)+'/'+str(tamanho)+' - '+j, end="\r")
                #if i > 10:
                #    break
                if j=="" or j=='-':
                    a,b = [],[]
                    valores[0].append(a)
                    valores[1].append(b)
                elif j=='EQM':
                    a,b=['EQM'],['','','','','']
                    valores[0].append(a)
                    valores[1].append(b)
                else:
                    id_projeto, titulo, statusprj, statususuario, statushektor = consulta_projeto(j)
                    statusceite, obsaceite, serial = consulta_pasta(id_projeto)
                    a=[titulo]
                    b=[statusprj,statususuario,statusceite,statushektor,serial,obsaceite]
                    valores[0].append(a)
                    valores[1].append(b)
                print("\033[K", str(i)+'/'+str(tamanho)+' - '+j, b, end="\r")
            #print(valores)
            #break
            progressopasta.value = 1
            porcentagempasta.value =text='{:.2f}%'.format(progressopasta.value*100)
            infopastas.value = f'Salvando status das pastas {aba.title}'
            progressopasta.update()
            porcentagempasta.update()
            infopastas.update()

            #sh.worksheet(aba.title).update("D2:D",valores[0])
            sh.worksheet(aba.title).update("H2:M",valores[1])

    print('\n' + hora_atual() + ': Pastas atualizadas!')
    infopastas.value = f'Pastas atualizadas!\n'
    infopastas.update()

if __name__ == '__main__':
    a=consulta_projeto('B-1131975')
    print(consulta_pasta(a[0]))
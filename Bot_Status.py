import os
import pandas as pd
import requests, gspread

from datetime import datetime


path = os.path.dirname(os.path.abspath(__file__))
gs = gspread.service_account(filename=os.path.join(path, 'service_account.json'))


# Chamada p/ atualizar planilha
planilha = gs.open_by_key('1KwjWPYpIem8o3DcLCloPWcUMEz0rnXrVjWjBCapf21M')
aba = planilha.worksheet('AIC')
projetos = aba.col_values(2)[1:]

# Variáveis

status = ''
status_p = ''

# Matrizes
status_obras = []; status_termos = []; datas_zps = []; status_projeto = [];
data0_pasta = []; data1_pasta = []; data2_pasta = []; data3_pasta = []
status0_pasta = []; status1_pasta = []; status2_pasta = []; status3_pasta = []
nome_responsavel = []; data_responsavel = []; data_responsavel2 = []; data_baixa = []; data_responsavel_baixa = []

for i in projetos:
    #------------------- Requisição dos dados dos projetos no GEOEX
    fim = False

    url = 'https://geoex.com.br/api/Programacao/ConsultarProjeto/Item'

    header = {
        'cookie': 'access_token=CfDJ8HtNYPpOjypBrd5pW1lM_mjnhtCrIlEAIezvBJa8tojZLzn-DM8l8GKvgaak455CfyIMWXPS0jkmAFQ7dkpKUwdYZQZqcJJ3jGq5-A1F6w6YkquMw7AKKGL1dyK_m6BhutGL33kijanwWapFFfVE3ZZHEw16_NvW2IyPsfy8BbsM5k9rEZFyIWpFgEYSnuQOn1URJIY56nosmMh5zLr_JoKcB_FrWisjgzKneLJoXNjixlgm5ZigEOHKByoxXexTvYuqUwcvpt-1retNQlzB6YHUBEu7jKkCBtUUH7Rf0ESbY_UB6qXTh02dnrfURwXErbaiM0fkz4e4a_5krKvNebmDaZyTXC4Bo0f6s04lKERPNPbW0WGnrpdEaFgp54mULLU2RXA7XpDi_ABPgWTZiUmcrXu6mh7D7bGuiIjsF1FTN1UxsDF1leHOmjiOw7xKe8APvY3Rwr37nXXTAB2lggU; _ga=GA1.1.2090438642.1705578696; TemaEscuro=true; FirmaId=1; BaseConhecimento.Informativos.ModoLista=false; BaseConhecimento.Tutoriais.ModoLista=false; ConsultarNota.Numero=9102044091; _ga_ZBQMHFHTL8=GS1.1.1720102407.266.1.1720102597.0.0.0; CadastroDemanda.Serial=CDM0007891%2F24; Home.Buscar.Texto=; BaseConhecimento.TutoriaisTerceiros.ModoLista=false; ConsultarProjeto.Numero=B-SIR3401; .AspNetCore.Session=CfDJ8HtNYPpOjypBrd5pW1lM%2FmgQQT6J7FgSh5kNU828%2Fu0koUwDmDvujmhstAVIQPwGDLbwR73uC1tKw8Cl3kXxYdC%2FmEsZNuBw53H%2Bp65Qn1l36aMn0Uy%2F55XkKhbSf0f7JBMb5Qo4h5ZZfaD71Xbrh%2FXi%2F3RmApfnmK1TkyDiCxID; cf_clearance=ojNrn3vk4AgorTEHWA_RB3gpQFwlhIhoEPgBIV5eSCs-1720607371-1.0.1.1-WGXQCc4lKQaPYyJG8u.sE79LRr5JqjBeqSuAFZ7bmDAMw3pI4eId24zb58tePTfFHt9cQ0.1MKTU2Z98m1jM9g',
        'Gxsessao': 'KGBLXTdvREQiXShgYCJuSygoIm9uJ10ibmBgS10mRC1gV3RuH1cmNDdXdHR0IjctblciNDRERCImV29vIidvVyc3b1dEN0RdNw==',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    }

    body = {
        'id': i
    }

    while True:
        try:
            r = requests.post(url = url, headers = header, json = body)
            if r.status_code!=200:
                print(i+' Erro na requisição: Code: '+str(r.status_code)+', Reason: '+str(r.reason), end='\r')
                fim = True
                continue
            if fim:
                fim = False
                print('')
            r = r.json()
            break
        except Exception as e:
            print('Exception', e)
            print('Requisição', r)
            print('Projeto ', i)
            raise Exception('Não foi possível acessar a página do GEOEX.')


    if r['Content'] != None and r['Content']['ProjetoId'] != None:

        id_projeto = r['Content']['ProjetoId']

        # Obtendo status das obras
        if r['Content']['StatusUsuario'] != None:
            status = r['Content']['StatusUsuario']['Descricao']
            status_obras.append([status])

        else:
            status_obras.append([''])


        # Obtendo status do projeto
        if r['Content']['StatusProjeto'] != None:
            status_p = r['Content']['StatusProjeto']['Descricao']
            status_projeto.append([status_p])

        else:
            status_projeto.append([''])


        # Obtendo status dos termos
        if r['Content']['GseProjeto'] != None:
            termo = r['Content']['GseProjeto']['Status']['Nome']
            status_termos.append([termo])

        else:
            status_termos.append([''])

        # Obtendo data da ZPS09
        if r['Content']['DtZps09'] != None:
            data = datetime.strptime(r['Content']['DtZps09'], '%Y-%m-%dT%H:%M:%S.%f%z').strftime('%d/%m/%Y')
            datas_zps.append([data])

        else:
            datas_zps.append([''])

        #------------------- Requisição dos status das pastas no GEOEX

        url = 'https://geoex.com.br/api/ConsultarProjeto/EnvioPasta/Itens'

        body = {
            'ProjetoId': id_projeto,
            'Paginacao': {
                'Pagina': '1',
                'TotalPorPagina': '10'
            }
        }

        r = requests.post(url, headers = header, json = body).json()
        print(r)
        n = r['Content']['Total']

        try:
            pasta0 = int(r['Content']['Items'][0].get('HistoricoStatusId', '')) 
            if pasta0 == 30:
                pasta0 = 'ACEITO'
            elif pasta0 == 1:
                pasta0 = 'CRIADO'
            elif pasta0 == 31:
                pasta0 = 'ACEITO COM RESTRIÇÕES'
            elif pasta0 == 32:
                pasta0 = 'REJEITADO'
            elif pasta0 == 6:
                pasta0 = 'CANCELADO'
            else:
                pasta0 = ''

            status0_pasta.append([pasta0]);

        except:
            pasta0 = ''; status0_pasta.append([pasta0]);
        
        try:
            nome_user = r['Content']['Items'][0]['UsuarioResponsavel']['Nome']
            nome_responsavel.append([nome_user]);
        except:
            nome_user = ''; nome_responsavel.append([nome_user]);
        
        try:
            data_user = r['Content']['Items'][0]['Data']
            data_user = datetime.fromisoformat(data_user)
            data_user = data_user.strftime("%d/%m/%y %H:%M")
            data_responsavel.append([data_user]);
        except:
            data_user = ''; data_responsavel.append([data_user]);

        try:
            data_user_responsavel = r['Content']['Items'][0]['DataResponsavel']
            data_user_responsavel = datetime.fromisoformat(data_user_responsavel)
            data_user_responsavel = data_user_responsavel.strftime("%d/%m/%y %H:%M")
            data_responsavel2.append([data_user_responsavel]);
        except:
            data_user_responsavel = ''; data_responsavel2.append([data_user_responsavel]);
            
        try:
            data_baixa = r['Content']['Items'][0]['DataBaixa']
            data_baixa = datetime.fromisoformat(data_baixa)
            data_baixa = data_baixa.strftime("%d/%m/%y %H:%M")
            data_responsavel_baixa.append([data_baixa]);
        except:
            data_baixa = ''; data_responsavel_baixa.append([data_baixa]);

    else:
        print('Não foi possível acessar o projeto', i, 'no GEOEX.')

        if r['IsUnauthorized']:
            print(r)
            raise Exception('Cookie inválido! Não autorizado')
        
        status_obras.append([]); status_termos.append([]); datas_zps.append([])
        status0_pasta.append([]); status_projeto.append([]);

    print('\nObtendo informações do projeto', i,
          '\nstatus_projeto: ', status, '/',
          '\nProjeto ID: ', id_projeto, '/', 'No. de pastas:', n,
          '\nstatus_pasta: ', pasta0,'\nstatus_projeto: ', status_p,
          '\nData User: ', data_user, '/', 
          '\nData Responsável: ', data_user_responsavel, '/',
          '\nData Baixa: ', data_baixa, '/',
          '\nNome Responsavel: ', nome_user, '/', 
          '\nData Responsavel: ', data_user, '/',)

dados_geral = {
    'status_projeto': [item[0] for item in status_projeto],
    'status_obras': [item[0] for item in status_obras],
    'status_termos': [item[0] for item in status_termos],
    'datas_zps': [item[0] for item in datas_zps],
    'status0_pasta': [item[0] for item in status0_pasta],
    'data_responsavel': [item[0] for item in data_responsavel],
    'nome_responsavel': [item[0] for item in nome_responsavel],
    'data_responsavel2': [item[0] for item in data_responsavel2],
    'data_responsavel_baixa': [item[0] for item in data_responsavel_baixa]
}

dados_df = pd.DataFrame(dados_geral)

print("\nAtualizando status dos projetos ..")
aba.update(range_name='C2:L', values=dados_df.values.tolist())
print("Status dos projetos atualizados!\n")
import flet as ft
from Consulta_Geoex_Bib_Flet import *
from time import time,sleep

#importar links das planilhas
planilhas = abre('_internal/meses.json')
meses = [a for a in planilhas]
marcados = []

#importar valores do geoex
data = abre('_internal/cookie.json')
cookie = data['cookie']
gxsessao = data['gxsessao']
useragent = data['useragent']

atualizando = False
atualizando1 = False
umavez = None
mesesselecionados = []
mesesdel = []
colunameses = []
colunamesesdel = []
parar = False

def main(page: ft.Page):
    global umavez, mesesselecionados, parar, mesesdel, colunameses, colunamesesdel

    page.title = "BOT PARA ATUALIZAR STATUS DAS MEDIÇÕES DO GEOEX"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.auto_scroll = True
    page.window_width = 800
    #page.window_height=500
    #page.theme_mode='light'
    page.theme = ft.Theme(color_scheme_seed='teal')

    conexao = ft.Text('-----', color=ft.colors.WHITE38, size=20, col={'sm': 3})
    cookienovo = ft.TextField(label='Cookie')
    gxsessaonovo = ft.TextField(label='Gxsessao')
    useragentnovo = ft.TextField(label='User-Agent')
    temponovo = ft.TextField(label='Tempo', value='0')
    progresso = ft.ProgressBar(value=0)
    porcentagem = ft.Text('--%', text_align=ft.TextAlign.CENTER)
    progressopasta = ft.ProgressBar(value=0)
    porcentagempasta = ft.Text('--%', text_align=ft.TextAlign.CENTER)
    nomeplanilha = ft.Text('Selecione a opção de atualização.', text_align=ft.TextAlign.CENTER, theme_style=ft.TextThemeStyle.HEADLINE_MEDIUM)
    popup = ft.AlertDialog()
    projetonovo = ft.TextField(label='Projeto')
    medicaonovo = ft.TextField(label='Medição')
    resultado = ft.Text('', theme_style=ft.TextThemeStyle.BODY_MEDIUM)
    mesnovo = ft.TextField(label='Nome do Mês', value='')
    linknovo = ft.TextField(label='Chave da Planilha', value='')
    intervalonovo = ft.TextField(label='Intervalo para salvar os valores na planilha', value='')

    def abrepopup(menssagem):
        popup.title = ft.Text(menssagem, text_align=ft.TextAlign.CENTER)
        page.dialog = popup
        popup.open = True
        page.update()

    def itens(vetor):
        item = []
        for i in list(reversed(vetor)):
            item.append(ft.Checkbox(label=i, value=False))
        return item

    def delitens(vetor):
        item = []

        def criar_apaga_mes(id):
            return lambda _: apaga_mes(id)
        
        for i in list(vetor):
            item.append(ft.ElevatedButton(text=i, on_click=criar_apaga_mes(i)))
            
        return item

    def sair(e):
        page.window_destroy()

    def checkgeoex(e):
        try:
            procura_projeto('B-1119157', cookie, gxsessao, useragent)
            conexao.color = ft.colors.GREEN
            conexao.value = 'Conectado'
            print('Conectado')
            conexao.update()
        except Exception as ex:
            print(ex)
            conexao.color = ft.colors.RED
            conexao.value = 'Desconectado'
            conexao.update()
            pass

    def close_dlg(janela):
        janela.open = False
        page.update()
    
    def atualizacookie(e):
        global cookie, gxsessao, useragent, data

        if cookienovo.value=='' and gxsessaonovo.value=='' and useragentnovo.value=='':
            return
        
        try:
            if cookienovo.value!='':
                cookie=str(cookienovo.value).strip()
            if gxsessaonovo.value!='':
                gxsessao=str(gxsessaonovo.value).strip()
            if useragentnovo.value!='':
                useragent=str(useragentnovo.value).strip()

            procura_projeto('B-1119157', cookie, gxsessao, useragent)

            data['cookie'], data['gxsessao'], data['useragent'] = cookie, gxsessao, useragent
            escreve_json('_internal/cookie.json', data)
            message = 'Cookie atualizado!'
            print(message)
            conexao.color = ft.colors.GREEN
            conexao.value = 'Conectado'
            conexao.update()
        except Exception as e:
            cookie, gxsessao, useragent = data['cookie'], data['gxsessao'], data['useragent']
            message = 'Não foi possível acessar o Geoex.'

        popup = ft.AlertDialog(title=ft.Text(message))
        page.dialog = popup
        popup.open = True
        page.update()

    def open_dlg(janela):
        page.dialog = janela
        janela.open = True
        page.update()
        janela.update()

    def atualizasemtemp():
        global atualizando1, mesesselecionados
        print('Atualizando sem temporização')

        if atualizando1:
            print('Já atualizando')
            abrepopup('Já atualizando')
            return
        
        atualizando1 = True
        print('atualizando')

        for i,j in enumerate(list(reversed(mesesselecionados))):
            print(j.value)
            if j.value:
                try:
                    atualiza_planilha(planilhas[meses[i]][0], progresso, porcentagem, nomeplanilha, planilhas[meses[i]][1])
                except Exception as e:
                    atualizando1 = False
                    print(e)
                    break
        
        temporestante.value = 'Atualizado!'
        temporestante.update()

        atualizando1 = False
        abrepopup('Planilha atualizada com sucesso!')

    def atualizatemp():
        global atualizando, mesesselecionados, parar
        print('Atualizando com temporização')#,parar, atualizando)
        
        if atualizando:
            print('Já atualizando')
            abrepopup('Já atualizando')
            return
        
        parar, atualizando = False, True
        print(parar, atualizando)
        open_dlg(janelatempo)

        while janelatempo.open==True:
            continue
        
        tempo = float(temponovo.value)
        
        while True:
            if parar:
                print('interrompido')
                parar, atualizando = False, False
                temporestante.value = 'Tempo até a próxima atualização: Interrompido'
                temporestante.update()
                break

            if tempo == 0:
                print('Tempo zerado')
                parar, atualizando = False, False
                temporestante.value = 'Tempo até a próxima atualização: 00:00 min'
                temporestante.update()
                break
            
            for i,j in enumerate(list(reversed(mesesselecionados))):
                print(j.value)
                if j.value:
                    try:
                        atualiza_planilha(planilhas[meses[i]][0], progresso, porcentagem, nomeplanilha, planilhas[meses[i]][1])
                    except Exception as e:
                        parar, atualizando = True, False
                        print(e)
                        print('Erro na atualização')
                        break
                    #print(parar)
            
            inicio = time()
            t = tempo
            
            while t > 0:
                if parar: break

                fim = time()
                t = tempo*60 - (fim - inicio)
                mins, secs = divmod(t, 60)
                timer = '{:02.0f}:{:02.0f}'.format(mins, secs)
                temporestante.value = 'Tempo até a próxima atualização: ' + timer + ' min'
                temporestante.update()
                print('Tempo ' + timer + ' min', end="\r")
                sleep(1)
    
    def try_atualizatemp(e):
        global parar, atualizando
        try:
            temporestante.value = 'Atualizando...'
            temporestante.update()
            atualizatemp()
        except Exception as e:
            parar, atualizando = True, False
            print(e)
            print('Erro na atualização')
            abrepopup('Erro na atualização\n'+str(e))
            temporestante.value = 'Tempo até a próxima atualização: Erro'
            temporestante.update()
            checkgeoex(e)

    def try_atualizasemtemp(e):
        global parar, atualizando
        try:
            temporestante.value = 'Atualizando...'
            temporestante.update()
            atualizasemtemp()
        except Exception as e:
            parar, atualizando = True, False
            print(e)
            print('Erro na atualização')
            abrepopup('Erro na atualização\n'+str(e))
            temporestante.value = 'Tempo até a próxima atualização: Erro'
            temporestante.update()
            checkgeoex(e)
    
    def try_atualizapasta(e):
        try:
            infopastas.value = 'Atualizando...'
            infopastas.update()
            atualiza_pasta(infopastas, progressopasta, porcentagempasta)
        except Exception as e:
            print(e)
            print('Erro na atualização')
            abrepopup('Erro na atualização\n'+str(e))
            infopastas.value = 'Erro'
            infopastas.update()
            checkgeoex(e)

    def paratemp(e):
        global parar
        parar = True

    def consulta(projeto, medicao):
        resultado.value = 'Buscando...'
        resultado.update()
        r = consulta_medicao_geoex(projeto, medicao, cookie, gxsessao, useragent)
        if r!=['','','','']:
            resultado.value = f'Status: {r[0]}\nData Postagem: {r[1]}\nData Atesto: {r[2]}\nData Validação: {r[3]}'
        else:
            resultado.value = 'Erro'
        resultado.update()

    mesesselecionados = itens(meses)
    mesesdel = delitens(meses)

    def salva_mes(mes, link, intervalo):
        if mes == '' and link == '' and intervalo == '':
            close_dlg(janelanovomes)
            abrepopup('Insira nome e link válido')
            return

        global planilhas, meses, mesesselecionados, mesesdel
        planilhas.update({mes:[link,intervalo]})
        escreve_json('_internal/meses.json', planilhas)
        meses = [a for a in planilhas]
        mesesselecionados.insert(0, ft.Checkbox(label=meses[-1], value=False))
        mesesdel.append(ft.ElevatedButton(text=meses[-1], on_click=lambda _: apaga_mes(meses[-1])))
        colunameses.update()
        page.update()
        close_dlg(janelanovomes)
        abrepopup('Planilha Adicionada')

    def apaga_mes(id):
        global mesesdel, mesesselecionados, colunameses, colunamesesdel
        print(meses)
        print(mesesdel)
        print(meses.index(id), id, len(mesesselecionados) - meses.index(id) - 1)
        planilhas.pop(id)
        mesesdel.pop(meses.index(id))
        mesesselecionados.pop(len(mesesselecionados) - meses.index(id) - 1)
        meses.pop(meses.index(id))

        colunamesesdel.update()
        colunameses.update()
        janelaapagames.update()
        escreve_json('_internal/meses.json', planilhas)

    umavez = ft.ElevatedButton(text='Atualizar uma vez', on_click=try_atualizasemtemp, col={'sm': 4})
    temporizado = ft.ElevatedButton(text='Atualização temporizada', on_click=try_atualizatemp, col={'sm': 4})
    para = ft.ElevatedButton(text='Parar temporizada', on_click=paratemp, col={'sm': 4})
    temporestante = ft.Text(theme_style=ft.TextThemeStyle.HEADLINE_MEDIUM)
    colunameses = ft.Column(wrap=True, spacing=10, run_spacing=10, controls=mesesselecionados, scroll=ft.ScrollMode.AUTO)
    colunamesesdel = ft.Column(wrap=True, spacing=10, run_spacing=10, controls=mesesdel, scroll=ft.ScrollMode.AUTO)
    
    atualizapastas = ft.ElevatedButton(text='Atualizar uma vez', on_click=try_atualizapasta)
    infopastas = ft.Text('Atualize os Status das Pastas', theme_style=ft.TextThemeStyle.HEADLINE_MEDIUM)

    janelageo = ft.AlertDialog(
        title=ft.Text('Dados Geoex'),
        content=ft.Column(
            controls=[
                ft.Text('Insira os dados'),
                cookienovo,
                gxsessaonovo,
                useragentnovo
            ],
            tight=True
        ),
        actions=[
            ft.TextButton('Atualizar', on_click=atualizacookie),
            ft.TextButton('Fechar', on_click=lambda _: close_dlg(janelageo))
        ],
        adaptive=True
    )
    
    janelatempo = ft.AlertDialog(
        title=ft.Text('Dados Geoex'),
        content=ft.Column(
            controls=[
                ft.Text('Insira o tempo'),
                temponovo
            ],
            tight=True
        ),
        actions=[
            ft.TextButton('Salvar', on_click=lambda _: close_dlg(janelatempo))
        ]
    )

    janelamedicao = ft.AlertDialog(
        title=ft.Text('Consultar Medição'),
        content=ft.Column(
            controls=[
                ft.Text('Insira os dados'),
                projetonovo,
                medicaonovo,
                ft.Container(content=resultado)
            ],
            tight=True
        ),
        actions=[
            ft.TextButton('Consultar', on_click=lambda _: consulta(projetonovo.value, medicaonovo.value)),
            ft.TextButton('Fechar', on_click=lambda _: close_dlg(janelamedicao))
        ],
        adaptive=True
    )

    janelanovomes = ft.AlertDialog(
        title=ft.Text('Adicionar Planilha'),
        content=ft.Column(
            controls=[
                ft.Text('Insira os dados'),
                mesnovo,
                linknovo,
                intervalonovo
            ],
            tight=True
        ),
        actions=[
            ft.TextButton('Salvar', on_click=lambda _: salva_mes(mesnovo.value, linknovo.value, intervalonovo.value)),
            ft.TextButton('Fechar', on_click=lambda _: close_dlg(janelanovomes))
        ],
        adaptive=True
    )

    janelaapagames = ft.AlertDialog(
        title=ft.Text('Remover Planilha'),
        content=ft.Column(
            controls=[
                ft.Text('Selecione a Planilha para Remover'),
                colunamesesdel
            ],
            tight=True
        ),
        actions=[
            ft.TextButton('Fechar', on_click=lambda _: close_dlg(janelaapagames))
        ],
        adaptive=True
    )

    botoes = ft.Container(col={'sm': 9}, content=ft.Column(spacing=10, controls=[
        ft.Container(content=ft.ResponsiveRow([
            ft.ElevatedButton(text="Consultar Medição", on_click=lambda _: open_dlg(janelamedicao), col={'sm': 6}),
            ft.ElevatedButton(text='Atualizar Valores Geoex', on_click=lambda _: open_dlg(janelageo), col={'sm': 6})
        ])),
        ft.Container(content=ft.ResponsiveRow([
            ft.ElevatedButton(text="Adicionar Planilha", on_click=lambda _: open_dlg(janelanovomes), col={'sm': 6}),
            ft.ElevatedButton(text="Remover Planilha", on_click=lambda _: open_dlg(janelaapagames), col={'sm': 6})
        ])),
        ft.Container(content=ft.ResponsiveRow([
            ft.ElevatedButton(text="Sair", on_click=sair, col={'sm': 12})
        ]))
    ]))
    
    menu = ft.Column(controls=[
        ft.Container(
            content=ft.Text('BOT PARA ATUALIZAR STATUS DAS MEDIÇÕES DO GEOEX', size=20, text_align=ft.TextAlign.CENTER),
            bgcolor=ft.colors.TEAL,
            border_radius=15,
            padding=15,
            alignment = ft.alignment.center
        ),
        ft.ResponsiveRow([
            ft.Text('Status Geoex: ', size=20, col={'sm': 3}),
            conexao,
            ft.IconButton(icon=ft.icons.UPDATE, icon_size=20, col={'sm': 1}, on_click=checkgeoex),
            ft.Text('By: Heli Neres', size=20, text_align=ft.TextAlign.END, col={'sm': 5})
        ]),
        ft.Divider(),
        botoes
    ])

    abamedicoes = [
        ft.Container(content=nomeplanilha, padding=15),
        ft.ResponsiveRow([
            ft.Container(
                content=ft.Column(controls=[
                    ft.Container(content=ft.ResponsiveRow([
                        umavez,
                        temporizado,
                        para
                    ]), padding=15),
                    ft.Container(content=ft.ResponsiveRow([
                        ft.Container(content=progresso, alignment=ft.alignment.bottom_center, col={'sm': 10}, padding=10),
                        ft.Container(content=porcentagem, alignment=ft.alignment.top_center, col={'sm': 2})
                    ]), padding=15),
                    temporestante
                ]),
                col={'sm': 9}),
            ft.Container(
                content=colunameses,
                border=ft.border.all(width=2, color=ft.colors.TEAL),
                border_radius=15,
                padding=5,
                col={'sm': 3}
            )
        ])
    ]

    abapastas = [
        ft.Container(content=infopastas, padding=15),
        ft.ResponsiveRow([
            ft.Column(controls=[
                ft.Container(content=ft.ResponsiveRow([
                    atualizapastas
                ])),
                ft.Container(content=ft.ResponsiveRow([
                    ft.Container(content=progressopasta, alignment=ft.alignment.bottom_center, col={'sm': 10}, padding=10),
                    ft.Container(content=porcentagempasta, alignment=ft.alignment.top_center, col={'sm': 2})
                ]), padding=30)
            ])
        ])
    ]

    t= ft.Tabs(
        selected_index=0,
        tabs=[
            ft.Tab(
                text='Menu',
                content=menu
            ),
            ft.Tab(
                text='Medições',
                content=ft.Column(
                    controls=abamedicoes
                )
            ),
            ft.Tab(
                text='Pastas',
                content=ft.Column(
                    controls=abapastas
                )
            )
        ]
    )

    page.add(t)
    page.update()
    
    checkgeoex(page)

if __name__ == '__main__':
    #ft.app(port=8550, target=main)
    ft.app(port=80 ,target=main, view=ft.AppView.WEB_BROWSER)
# -*- coding: utf-8 -*-
import datetime, json, sys, io, os, traceback, PySimpleGUI as sg
from threading import Thread
from PIL import Image, ImageDraw
from PySimpleGUI import BUTTON_TYPE_READ_FORM, FILE_TYPES_ALL_FILES, theme_background_color, theme_button_color, BUTTON_TYPE_BROWSE_FILE, BUTTON_TYPE_BROWSE_FOLDER
from base64 import b64encode

import comum, analisa

nome_da_janela = 'Analisa Arquivos IRPF'
curvatura_do_botao = 0.8
margem_da_janela = (25, 25)
tamanho_padrao = (750, 400)

icone = 'Assets/AI_icon.ico'
exclamacao = 'Assets/alert_ponto-de-exclamacao.png'
exclamacao_soft = 'Assets/alert_ponto-de-exclamacao_soft.png'
verificacao = 'Assets/alert_verificado.png'
alerta_erro = 'Assets/alert_cancelar.png'
dados_modo = 'Assets/modo.txt'
controle_rotinas = 'Log/Controle.txt'
controle_botoes = 'Log/Buttons.txt'
dados_elementos = 'Log/window_values.json'
# Define o ícone global da aplicação
sg.set_global_icon(icone)


def RoundedButton(button_text=' ', corner_radius=0.1, button_type=BUTTON_TYPE_READ_FORM, target=(None, None), tooltip=None, file_types=FILE_TYPES_ALL_FILES, initial_folder=None, default_extension='',
                  disabled=False, change_submits=False, enable_events=False, image_size=(None, None), image_subsample=None, border_width=0, size=(None, None), auto_size_button=None,
                  button_color=None, background_color=False, disabled_button_color=None, highlight_colors=None, mouseover_colors=(None, None), use_ttk_buttons=None,
                  font=None, bind_return_key=False, focus=False, pad=None, key=None, right_click_menu=None, expand_x=False, expand_y=False, visible=True, metadata=None):
    if not background_color: background_color = theme_background_color()
    
    if None in size:
        multi = 5
        size = (((len(button_text) if size[0] is None else size[0]) * 5 + 20) * multi,
                20 * multi if size[1] is None else size[1])
    if button_color is None:
        button_color = theme_button_color()
    btn_img = Image.new('RGBA', size, (0, 0, 0, 0))
    corner_radius = int(corner_radius / 2 * min(size))
    poly_coords = (
        (corner_radius, 0),
        (size[0] - corner_radius, 0),
        (size[0], corner_radius),
        (size[0], size[1] - corner_radius),
        (size[0] - corner_radius, size[1]),
        (corner_radius, size[1]),
        (0, size[1] - corner_radius),
        (0, corner_radius),
    )
    pie_coords = [
        [(size[0] - corner_radius * 2, size[1] - corner_radius * 2, size[0], size[1]),
         [0, 90]],
        [(0, size[1] - corner_radius * 2, corner_radius * 2, size[1]), [90, 180]],
        [(0, 0, corner_radius * 2, corner_radius * 2), [180, 270]],
        [(size[0] - corner_radius * 2, 0, size[0], corner_radius * 2), [270, 360]],
    ]
    brush = ImageDraw.Draw(btn_img)
    brush.polygon(poly_coords, button_color[1])
    for coord in pie_coords:
        brush.pieslice(coord[0], coord[1][0], coord[1][1], button_color[1])
    data = io.BytesIO()
    btn_img.thumbnail((size[0] // 3, size[1] // 3), resample=Image.LANCZOS)
    btn_img.save(data, format='png', quality=95)
    btn_img = b64encode(data.getvalue())
    return sg.Button(button_text=button_text, button_type=button_type, target=target, tooltip=tooltip,
                     file_types=file_types, initial_folder=initial_folder, default_extension=default_extension,
                     disabled=disabled, change_submits=change_submits, enable_events=enable_events,
                     image_data=btn_img, image_size=image_size,
                     image_subsample=image_subsample, border_width=border_width, size=size,
                     auto_size_button=auto_size_button, button_color=(button_color[0], background_color),
                     disabled_button_color=disabled_button_color, highlight_colors=highlight_colors,
                     mouseover_colors=mouseover_colors, use_ttk_buttons=use_ttk_buttons, font=font,
                     bind_return_key=bind_return_key, focus=focus, pad=pad, key=key, right_click_menu=right_click_menu,
                     expand_x=expand_x, expand_y=expand_y, visible=visible, metadata=metadata)


def show_alert(text, image_path='', buttons=["OK"], no_titlebar=True):
    f = open(dados_modo, 'r', encoding='utf-8')
    modo = f.read()
    
    if modo == 'claro':
        cor_de_fundo = '#EFEFEF'
        cor_do_texto = '#000000'
    else:
        cor_de_fundo = '#0F0F0F'
        cor_do_texto = '#F8F8F8'
    
    layout = [
        [sg.Image(image_path, background_color=cor_de_fundo), sg.Text('', background_color=cor_de_fundo), sg.Text(text, size=(50, None), text_color=cor_do_texto, auto_size_text=True, expand_y=True, expand_x=True, background_color=cor_de_fundo)],
        [sg.Text('', background_color=cor_de_fundo)]
    ]
    
    layout_2 = [[layout], [sg.Column([[sg.Text('', background_color=cor_de_fundo)]], expand_x=True, background_color=cor_de_fundo), sg.Column([[RoundedButton(button, corner_radius=curvatura_do_botao, button_color=(cor_de_fundo, '#848484'), background_color=cor_de_fundo) for button in buttons]], background_color=cor_de_fundo)]]
    
    window = sg.Window('Alerta', layout_2, modal=True, margins=margem_da_janela, no_titlebar=no_titlebar, keep_on_top=True, background_color=cor_de_fundo)
    
    event, values = window.read()
    window.close()
    return event


# controle de versão, exibe um alerta e atualiza o programa se o arquivo de controle de versão for diferente da versão descrita na variável a baixo
versao_atual = '1.0'
caminho_arquivo = f'T:\ROBÔ\_Executáveis\\{nome_da_janela}\\{nome_da_janela}.exe'

cr = open(f'T:\\ROBÔ\\_Executáveis\\{nome_da_janela}\\Versão.txt', 'r', encoding='utf-8').read()
if str(cr) != versao_atual:
    while True:
        atualizar = show_alert(image_path=exclamacao_soft, text='Existe uma nova versão do programa, deseja atualizar agora?', buttons=['Novidades da atualização', 'Sim', 'Não'], no_titlebar=False)
        if atualizar == 'Novidades da atualização':
            os.startfile(f'T:\\ROBÔ\\_Executáveis\\{nome_da_janela}\\Sobre.pdf')
        if atualizar == 'Sim':
            break
        elif atualizar == 'Não':
            break
    
    if atualizar == 'Sim':
        os.startfile(caminho_arquivo)
        sys.exit()


if __name__ == '__main__':
    # Função para salvar os valores dos elementos em um arquivo JSON
    def save_values(values, filename=dados_elementos):
        with open(filename, 'w', encoding='latin-1') as f:
            json.dump(values, f)
    
    
    # Função para carregar os valores dos elementos de um arquivo JSON
    def load_values(filename=dados_elementos):
        if os.path.exists(filename):
            with open(filename, 'r', encoding='latin-1') as f:
                return json.load(f)
        return {}
    
    
    def update_window_elements(window, updates):
        # Loop para aplicar as atualizações
        for key, update_args in updates.items():
            window[key].update(**update_args)
    

    # limpa o arquivo que salva o estado dos elementos preenchidos e o estado anterior do botão de abrir os resultados
    with open(dados_elementos, 'w', encoding='utf-8') as f:
        f.write('')
    with open(controle_botoes, 'w', encoding='utf-8') as f:
        f.write('')
    
    sg.LOOK_AND_FEEL_TABLE['tema_claro'] = {'BACKGROUND': '#F8F8F8',
                                            'TEXT': '#000000',
                                            'INPUT': '#F8F8F8',
                                            'TEXT_INPUT': '#000000',
                                            'SCROLL': '#F8F8F8',
                                            'BUTTON': ('#F8F8F8', '#F8F8F8'),
                                            'PROGRESS': ('#fca400', '#D7D7D7'),
                                            'BORDER': 0,
                                            'SLIDER_DEPTH': 0,
                                            'PROGRESS_DEPTH': 0, }
    
    sg.LOOK_AND_FEEL_TABLE['tema_escuro'] = {'BACKGROUND': '#000000',
                                             'TEXT': '#F8F8F8',
                                             'INPUT': '#000000',
                                             'TEXT_INPUT': '#F8F8F8',
                                             'SCROLL': '#000000',
                                             'BUTTON': ('#000000', '#000000'),
                                             'PROGRESS': ('#fca400', '#2A2A2A'),
                                             'BORDER': 0,
                                             'SLIDER_DEPTH': 0,
                                             'PROGRESS_DEPTH': 0}
    
    # define o tema baseado no tema que estava selecionado na última vêz que o programa foi fechado
    f = open(dados_modo, 'r', encoding='utf-8')
    modo = f.read()
    sg.theme(f'tema_{modo}')
    

    def janela_principal():
        def cria_layout():
            coluna_ui = [
                [sg.Button('AJUDA', key='-ajuda-'),
                 sg.Button('SOBRE', key='-sobre-'),
                 sg.Button('LOG DO SISTEMA', key='-log_sistema-', disabled=True),
                 sg.Text('', expand_x=True), sg.Text('', key='-relogio-'),
                 sg.Button('', key='-tema-', font=("Arial", 11), border_width=0),
                 ],
                [sg.Text('')],
                [sg.Text('')],
                
                [sg.pin(sg.Text(text='Analisando arquivos dessa pasta:', key='-pasta_arquivos_texto-', visible=False)),
                 sg.pin(RoundedButton('Selecione a pasta com os arquivos para analisar:', key='-pasta_arquivos-', corner_radius=curvatura_do_botao, button_color=(None, '#848484'),
                                      size=(920, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-input_dados-')),
                 sg.pin(sg.InputText(expand_x=True, key='-input_dados-', text_color='#fca400'), expand_x=True)],
                
                [sg.pin(sg.Text(text='Os resultados estão sendo salvos aqui:', key='-pasta_resultados_texto-', visible=False)),
                 sg.pin(RoundedButton('Selecione a pasta para salvar os resultados:', key='-pasta_resultados-', corner_radius=curvatura_do_botao, button_color=(None, '#848484'),
                                      size=(850, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-output_dir-')),
                 sg.pin(sg.InputText(expand_x=True, key='-output_dir-', text_color='#fca400'), expand_x=True)],
                [sg.Text('', expand_y=True)],
                
                [sg.Text('', key='-Mensagens_2-')],
                [sg.Text('', key='-Mensagens-')],
                [sg.Text(text='', key='-Progresso_texto-'),
                 sg.ProgressBar(max_value=0, orientation='h', size=(5, 5), key='-progressbar-', expand_x=True, visible=False)],
                [sg.pin(RoundedButton('INICIAR', key='-iniciar-', corner_radius=curvatura_do_botao, button_color=(None, '#fca400'))),
                 sg.pin(RoundedButton('ENCERRAR', key='-encerrar-', corner_radius=curvatura_do_botao, button_color=(None, '#fca400'), visible=False)),
                 sg.pin(RoundedButton('ABRIR RESULTADOS', key='-abrir_resultados-', corner_radius=curvatura_do_botao, button_color=(None, '#fca400'), visible=False))]
            ]
            
            coluna_terminal = [
                [RoundedButton(button_text='DESATIVAR TERMINAL', corner_radius=curvatura_do_botao, key='-dev_mode-', button_color=('black', 'red')),
                 sg.Button('🗑', key='-limpa_console-', font=("Arial", 11))],
                [sg.Output(expand_x=True, expand_y=True, key='-console-')]
            ]
            
            return [[sg.Column(coluna_ui, expand_y=True, expand_x=True), sg.Column(coluna_terminal, expand_y=True, expand_x=True, key='-terminal-', visible=False)]]
        
        # guarda a janela na variável para manipula-la
        return sg.Window(nome_da_janela + ' - Busca Herança e Meação', cria_layout(), return_keyboard_events=True, use_default_focus=False, resizable=True, finalize=True, margins=margem_da_janela)
    
    
    def captura_values():
        # recebe os valores dos inputs
        try:
            pasta_arquivos = values['-input_dados-']
            pasta_final = values['-output_dir-']
        except:
            pasta_arquivos = None
            pasta_final = None
            
        return pasta_arquivos, pasta_final
    
    
    def run_script_thread():
        # habilita e desabilita os botões conforme necessário
        updates = {
            '-pasta_arquivos-': {'visible': False},
            '-pasta_arquivos_texto-': {'visible': True},
            '-pasta_resultados-': {'visible': False},
            '-pasta_resultados_texto-': {'visible': True},
            '-iniciar-': {'visible': False},
            '-encerrar-': {'visible': True},
            '-abrir_resultados-': {'visible': True}
                }
        update_window_elements(window_principal, updates)

        # controle para saber se os botões estavam visíveis ao trocar o tema da janela
        with open(controle_botoes, 'w', encoding='utf-8') as f:
            f.write('visible')
        
        try:
            # Chama a função que executa o script
            analisa.executa_analisador(window_principal, pasta_arquivos, pasta_final)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            print(traceback_str)
            comum.escreve_doc(f'Traceback: {traceback_str}\n\n'
                              f'Erro: {erro}')
            window_principal['-log_sistema-'].update(disabled=False)
            window_principal.write_event_value('-ERRO-', 'alerta')
        
        updates = {
            '-pasta_arquivos-': {'visible': True},
            '-pasta_arquivos_texto-': {'visible': False},
            '-pasta_resultados-': {'visible': True},
            '-pasta_resultados_texto-': {'visible': False},
            '-iniciar-': {'visible': True},
            '-encerrar-': {'visible': False},
            '-Mensagens-': {'value': ''},
            '-Mensagens_2-': {'value': ''},
            '-Progresso_texto-': {'value': ''},
            '-progressbar-': {'visible': False},
                }
        update_window_elements(window_principal, updates)
        
        with open(controle_rotinas, 'w', encoding='utf-8') as f:
            f.write('STOP')
    
    
    # inicia as variáveis das janelas
    window_principal = janela_principal()
    # Definindo o tamanho mínimo da janela
    window_principal.set_min_size(tamanho_padrao)
    
    with open(controle_rotinas, 'w', encoding='utf-8') as f:
        f.write('STOP')
    
    while True:
        # Obtenha o widget Tkinter subjacente para o elemento Output
        output_widget = window_principal['-console-'].Widget
        
        # configura o tema do fundo dos botões conforme o tema que estava quando o programa foi fechado da última ves
        f = open(dados_modo, 'r', encoding='utf-8')
        modo = f.read()
        if modo == 'claro':
            for key in ['-ajuda-', '-sobre-', '-log_sistema-', '-limpa_console-']:
                window_principal[key].update(button_color=('#000000', None))
            window_principal['-tema-'].update(text='☀', button_color=('#FFC100', '#F8F8F8'))
            # Ajuste a borda diretamente no widget Tkinter
            output_widget.configure(highlightthickness=1, highlightbackground="black")
        else:
            for key in ['-ajuda-', '-sobre-', '-log_sistema-', '-limpa_console-']:
                window_principal[key].update(button_color=('#F8F8F8', None))
            window_principal['-tema-'].update(text='🌙', button_color=('#00C9FF', '#000000'))
            # Ajuste a borda diretamente no widget Tkinter
            output_widget.configure(highlightthickness=1, highlightbackground="white")
        
        # captura o evento e os valores armazenados na interface
        window, event, values = sg.read_all_windows(timeout=1000)
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            pasta_arquivos, pasta_final = captura_values()
        
        # salva o estado da interface
        save_values(values)
        
        # ↓↓↓ CHECKBOX ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        
        # controle de checkboxes
        checkboxes_0 = ['-opcao_1-', '-opcao_2-']
        if event in ('-opcao_1-', '-opcao_2-'):
            for checkbox_0 in checkboxes_0:
                if checkbox_0 != event:
                    window[checkbox_0].update(value=False)
                else:
                    continuar_rotina = checkbox_0
        
        
        # ↑↑↑ CHECKBOX ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
        
        if event == sg.WIN_CLOSED or event == "-Fechar-":
            if window == window_principal:  # if closing win 1, exit program
                break
        
        # troca o tema
        elif event == '-tema-':
            f = open(dados_modo, 'r', encoding='utf-8')
            modo = f.read()
            
            if modo == 'claro':
                sg.theme('tema_escuro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('escuro')
            else:
                sg.theme('tema_claro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('claro')
            
            window_principal.close()  # Fecha a janela atual
            # inicia as variáveis das janelas
            window_principal, window_configura = janela_principal(), None
            # Definindo o tamanho mínimo da janela
            window_principal.set_min_size(tamanho_padrao)
            # retorna os elementos preenchidos
            values = load_values()
            for key, value in values.items():
                if key in window_principal.AllKeysDict:
                    try:
                        window_principal[key].update(value)
                    except Exception as e:
                        print(f"Erro ao atualizar o elemento {key}: {e}")
            
            window_principal['-pasta_arquivos-'].update('Selecione a pasta com os arquivos para analisar:')
            window_principal['-pasta_resultados-'].update('Selecione a pasta para salvar os resultados:')
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                pasta_arquivos, pasta_final = captura_values()
            
            # recupera o estado do botão de abrir resultados
            cb = open(controle_botoes, 'r', encoding='utf-8').read()
            if cb == 'visible':
                window_principal['-abrir_resultados-'].update(visible=True)
        
        elif event == '-log_sistema-':
            os.startfile('Log')
        
        elif event == '-ajuda-':
            os.startfile(f'Docs\Manual do usuário - {nome_da_janela}.pdf')
        
        elif event == '-sobre-':
            os.startfile('Docs\Sobre.pdf')
        
        elif event == '-iniciar-':
            pasta_arquivos, pasta_final = captura_values()
            
            alerta_exibido = False
            
            # verifica se os inputs foram preenchidos corretamente
            for elemento in [(pasta_arquivos, 'Por favor informe a pasta com os arquivos para analisar.'), (pasta_final, 'Por favor informe a pasta onde será salvo os resultados.')]:
                if not elemento[0]:
                    show_alert(image_path=exclamacao, text=elemento[1])
                    alerta_exibido = True
                    break
            
            if not alerta_exibido:
                with open(controle_rotinas, 'w', encoding='utf-8') as f:
                    f.write('START')
                
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()
        
        elif event == '-encerrar-':
            window_principal['-Mensagens_2-'].update('')
            window_principal['-Mensagens-'].update('Encerrando, aguarde...')
            with open(controle_rotinas, 'w', encoding='utf-8') as f:
                f.write('ENCERRAR')
        
        elif event == '-abrir_resultados-':
            os.makedirs(pasta_final, exist_ok=True)
            try:
                os.startfile(pasta_final)
            except FileNotFoundError:
                show_alert(image_path=exclamacao_soft, text='O sistema não pode abrir essa pasta.')
        
        # ↓↓↓ ALERTAS ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        
        elif event == '-FINALIZADO-':
            show_alert(image_path=verificacao, text=f'Execução finalizada.')
        
        elif event == '-ENCERRADO-':
            show_alert(image_path=exclamacao, text=f'Execução encerrada.')
        
        elif event == '-ERRO-':
            show_alert(image_path=alerta_erro, text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
        
        elif event == '-ERROR_ZERO_BALANCE-':
            show_alert(image_path=exclamacao, text='Não consta saldo para resolução do captcha, mande um e-mail solicitando a recarga para o desenvolvedor.')

        # ↑↑↑ ALERTAS ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
        # ↓↓↓ BOTÕES DE AJUDA ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        
        elif event == '-ajuda_1-':
            show_alert(image_path=exclamacao_soft, text=f'Explicação da opção')
        
        # ↑↑↑ BOTÕES DE AJUDA ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
        # ↓↓↓ DEV MODE ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        
        elif event == 'Control_L:17':
            code = ''
        
        elif event == '-limpa_console-':
            window_principal['-console-'].update('')
        
        elif event == '-dev_mode-':
            window_principal['-terminal-'].update(visible=False)
            code = ''
        
        if event == 'Up:38' or event == 'Down:40' or event == 'Left:37' or event == 'Right:39' or event == 'a' or event == 'b':
            code += event
        else:
            code = ''
        
        if code == 'Up:38Up:38Down:40Down:40Left:37Right:39Left:37Right:39ab':
            window_principal['-terminal-'].update(visible=True)
            code = ''
        
        # ↑↑↑ DEV MODE ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
        # ↓↓↓ RELÓGIO ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        
        today = datetime.datetime.today()
        days = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
        months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        str_month = months[today.month - 1]  # obtemos o número do mês e subtraímos 1 para ter a correspondência correta com a nossa lista de meses
        str_weekday = days[today.weekday()]  # obtemos o número do dia da semana, neste caso o número 0 é segunda-feira e coincide com os indexes da nossa lista
        
        # código para a hora exata:  | %H:%M:%S
        window_principal['-relogio-'].update(f'{datetime.datetime.now().strftime(f"{str_weekday}, %d de {str_month} de %Y")}')
        
        # ↑↑↑ RELÓGIO ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
    
    window_principal.close()

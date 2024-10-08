import pandas as pd
import pyautogui
import time
import openpyxl
import math




def poste():

    for i, poste in enumerate(planilha_df['postes']):

        if math.isnan(float(i)):
            break

        numero = int(poste)
        numero_str = str(numero)
        coordenada_poste = planilha_df.loc[i, 'coordenadas']
        coordenada_estrutura = planilha_df.loc[i, 'coor estrutura']
        ater = str(planilha_df.loc[i, 'ater']).upper()
        poste_instalado = str(planilha_df.loc[i, 'poste instalado']).upper()
        poste_existente = str(planilha_df.loc[i, 'poste existente']).upper()
        poste_retirado = str(planilha_df.loc[i, 'poste retirado']).upper()
        troca_poste = str(planilha_df.loc[i, 'troca']).upper()
        estrutura_instalada = str(planilha_df.loc[i, 'estrutura instalada']).upper()
        estrutura_retirada = str(planilha_df.loc[i, 'estrutura retirada']).upper()
        chave_faca = str(planilha_df.loc[i, 'chave faca']).upper()
        numero_chave = str(planilha_df.loc[i, 'numero chave']).upper()
        cabo_instalado = planilha_df.loc[0, 'cabo instalado']
        cabo_retirado = planilha_df.loc[0, 'cabo retirado']
        vao = str(planilha_df.loc[i, 'vao']).upper()
        qtd_estai = str(planilha_df.loc[i, 'qtd est']).upper()
        bloco_poste = str(planilha_df.loc[i, 'bloco poste']).upper()

#poste simbolo

        if troca_poste == "X":
            print(f'Inserindo Poste do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write(f'{bloco_poste}')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        else:

            print(f'Inserindo Poste do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write(f'{bloco_poste}')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#indicacao de poste trocado

        if troca_poste == "X":
            print(f'Inserindo indicacao de poste trocado {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('pinst')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(f'{poste_instalado}')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('pret')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(f'{poste_retirado}')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        else:
            print(f'Inserindo indicacao de poste existente {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('t')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(f'{poste_existente}')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#estrutura

        print(f'Inserindo Estrutura do poste {numero}')

        if estrutura_instalada == "N4":
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        if estrutura_instalada == "TE":
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        if estrutura_instalada == "HT":
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('anco')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        pyautogui.moveTo(619, 390)
        pyautogui.click(619, 390)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('estinst')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(f'{estrutura_instalada}')
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.moveTo(619, 390)
        pyautogui.click(619, 390)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('estret')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(f'{estrutura_retirada}')
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)




#poste coordenada

        print(f'Inserindo coordenada do poste {numero}')
        pyautogui.moveTo(619, 390)
        pyautogui.click(619, 390)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('coord')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(f'{coordenada_poste}')
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)

#aterramento

        if ater == 'X':
            print(f'Inserindo Aterramento do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('aterpp')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#chave faca

        if chave_faca == 'X':
            print(f'Inserindo chave faca do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('cfcp')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('nchave')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(f'{numero_chave}')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#cabo
        print(f'Inserindo cabo do poste {numero}')
        pyautogui.moveTo(619, 390)
        pyautogui.click(619, 390)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('cabo')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(f'{cabo_instalado}')
        pyautogui.press('tab')
        pyautogui.write(f'{cabo_retirado}')
        pyautogui.press('tab')
        pyautogui.write(f'{vao}')
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)

#numeração

        print(f'Inserindo Numeração do poste {numero}')
        pyautogui.moveTo(545, 406)
        pyautogui.click(545, 406)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('num')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(numero_str)
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)

#estai
        print(f'Inserindo estai do poste {numero}')
        pyautogui.moveTo(545, 406)
        pyautogui.click(545, 406)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write(f'{qtd_estai}')
        pyautogui.press('enter')
        pyautogui.write(coordenada_poste)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)


def rede():
    print('Inserindo Rede')
    pyautogui.moveTo(80, 280)
    pyautogui.click(80, 280)
    pyautogui.click(80, 280)
    pyautogui.click(80, 280)
    pyautogui.write(coord_aglutinada, 0.25)
    pyautogui.press('esc')


def prancha():
    print('Inserindo Prancha')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('pranchaa1')
    pyautogui.press('enter')
    pyautogui.write(coor_prancha)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(cod_cidade)
    pyautogui.press('tab')
    pyautogui.write(cidade)
    pyautogui.press('tab')
    pyautogui.write(data_lev)
    pyautogui.press('tab')
    pyautogui.write(topografo)
    pyautogui.press('tab')
    pyautogui.write(data_projeto)
    pyautogui.press('tab')
    pyautogui.write(numero_ns)
    pyautogui.press('tab')
    pyautogui.write(alimentador)
    pyautogui.press('tab')
    pyautogui.write(cidade)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(tensao)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc')

#operação

print('AUTOPRO - AUTOMATIZADOR DE DESENHOS DE PROJETOS')
print('BEM VINDO AO MODULO RECONDUTORAMENTO')
print()

print()
print('             Orientações:')
print()
print('1 - Deixe a tela do Autocad aberta e maximizada.')
print()
print('2 - Desative o CAPS LOCK do seu teclado.')
print()
print('3 - Não faça nenhum intervenção durante o processo. Deixe o Robo trabalhar!')
print()
print()
print('Iniciando......')


pyautogui.PAUSE = 1.2
pyautogui.alert('Não use o computador enquanto o programa funciona!')

notas_df = pd.read_excel('notas.xlsx')

for i, notas in enumerate(notas_df['notas']):
    arquivo = notas_df.loc[i, 'notas']
    planilha_df = pd.read_excel(f'{arquivo}.xlsm', sheet_name='DADOS')

    coord_aglutinada = planilha_df.loc[0, 'aglutinadas']
    coord_aglutinada_sem_inicial = planilha_df.loc[0, 'aglutinadas menos inicial']
    ponto_inicial = str(planilha_df.loc[0, 'coordenadas'])
    alimentador = str(planilha_df.loc[0, 'alimentador']).upper()
    tensao = str(planilha_df.loc[0, 'tensao']).upper()
    num_ns = int(planilha_df.loc[0, 'ns'])
    numero_ns = str(num_ns)
    coor_prancha = str(planilha_df.loc[0, 'coor prancha'])
    cidade = str(planilha_df.loc[0, 'cidade']).upper()
    cod_cidade_num = int(planilha_df.loc[0, 'cod cidade'])
    cod_cidade = str(cod_cidade_num)
    data_lev = str(planilha_df.loc[0, 'data lev'])
    topografo = str(planilha_df.loc[0, 'topografo']).upper()
    data_projeto = str(planilha_df.loc[0, 'data proj'])

    print(f'Iniciando desenho da ns {arquivo}')
    poste()
    time.sleep(1)
    rede()
    time.sleep(1)
    prancha()
    print(f'Fim do desenho da ns {arquivo}')

print('Fim do Processo')
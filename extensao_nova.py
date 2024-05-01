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
        coordenada_numero = planilha_df.loc[i, 'coor numero']
        coordenada_dc = planilha_df.loc[i, 'coor dc']
        estrutura = str(planilha_df.loc[i, 'estrutura']).upper()
        tipo_poste = planilha_df.loc[i, 'tipo poste']
        atprov = str(planilha_df.loc[i, 'atprov']).upper()
        ater = str(planilha_df.loc[i, 'ater']).upper()
        prmt = str(planilha_df.loc[i, 'prmt']).upper()
        capacidade = planilha_df.loc[i, 'capacidade']
        equipamento = str(planilha_df.loc[i, 'equipamento']).upper()
        tipo_trafo = str(planilha_df.loc[0, 'trafo'])
        elo = str(planilha_df.loc[0, 'elo']).upper()
        elipse_chave = f'100A-1KA-{elo}'
        bloco_poste = str(planilha_df.loc[i, 'bloco poste']).upper()
        esforco = str(planilha_df.loc[i, 'esforco']).upper()
        resultante = str(planilha_df.loc[i, 'resultante']).upper()

#poste

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

#estrutura

        print(f'Inserindo Estrutura do poste {numero}')
        pyautogui.moveTo(619, 390)
        pyautogui.click(619, 390)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('eliest')
        pyautogui.press('enter')
        pyautogui.write(coordenada_estrutura)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(f'{estrutura}-{tipo_poste}')
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)

#aterramento

        if ater == 'X':
            print(f'Inserindo Aterramento e Aterramento Temporário do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('atatprovp')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#Para-raios mt

        if prmt == 'X':
            print(f'Inserindo Para-raios do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('prmtp')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#concretagem 600
        if capacidade == 600:
            print(f'Inserindo Diametro de concretagem do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('dc9')
            pyautogui.press('enter')
            pyautogui.write(coordenada_dc)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('concrec')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#concretagem 1000
        if capacidade == 000:
            print(f'Inserindo Diametro de concretagem do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('dc13')
            pyautogui.press('enter')
            pyautogui.write(coordenada_dc)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('concrec')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

# Esforco

        if esforco == 'X':
            print(f'Inserindo esforço do poste {numero}')
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('result')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(f'Rf = {resultante} daN')
            pyautogui.press('tab')
            pyautogui.write(f'Poste {numero}')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

#equipamento

        if equipamento == 'TRF 1':
            print(f'Inserindo Trafo do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('trafo')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(tipo_trafo)
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')
            time.sleep(1)

        if equipamento == 'CHV R':
            print(f'Inserindo chave repetidora do poste {numero}')
            pyautogui.moveTo(619, 390)
            pyautogui.click(619, 390)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('repetidora')
            pyautogui.press('enter')
            pyautogui.write(coordenada_poste)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('elichave')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.write(elipse_chave)
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')

        if equipamento == 'CHV FA':
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
            time.sleep(1)
            pyautogui.write('dd')
            pyautogui.press('enter')
            pyautogui.write('elichavefc')
            pyautogui.press('enter')
            pyautogui.write(coordenada_estrutura)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('tab')
            pyautogui.press('enter')
            pyautogui.press('esc')

#numeração

        print(f'Inserindo Numeração do poste {numero}')
        pyautogui.moveTo(545, 406)
        pyautogui.click(545, 406)
        pyautogui.write('dd')
        pyautogui.press('enter')
        pyautogui.write('num')
        pyautogui.press('enter')
        pyautogui.write(coordenada_numero)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.write(numero_str)
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        time.sleep(1)

def rede():
    print('Inserindo Rede')
    pyautogui.moveTo(98, 280)
    pyautogui.click(98, 280)
    pyautogui.click(98, 280)
    pyautogui.click(98, 280)
    pyautogui.click(98, 280)
    pyautogui.write(coord_aglutinada, 0.25)
    pyautogui.press('esc')

def estrada():
    print('Inserindo Estrada')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('ml')
    time.sleep(1)
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial)
    pyautogui.press('enter')
    pyautogui.write(coord_aglutinada_sem_inicial, 0.25)
    pyautogui.press('esc')

def cliente():
    print('Inserindo Cliente')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('cliente')
    pyautogui.press('enter')
    pyautogui.write(ponto_trafo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(nome_cliente)
    pyautogui.press('tab')
    pyautogui.write(telefone_cliente)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('cliente2')
    pyautogui.press('enter')
    pyautogui.write(ponto_trafo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(disj_cliente)
    pyautogui.press('tab')
    pyautogui.press('enter')
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
    pyautogui.write(nome_cliente)
    pyautogui.press('tab')
    pyautogui.write(cidade)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc')

#operação

print('AUTOPRO - AUTOMATIZADOR DE DESENHOS DE PROJETOS')
print('MODULO EXTENSÃO PELA ESTRADA')
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
    ponto_trafo = str(planilha_df.loc[0, 'pto trafo'])
    nome_cliente = str(planilha_df.loc[0, 'cliente']).upper()
    telefone_cliente = str(planilha_df.loc[0, 'telefone'])
    disj = str(planilha_df.loc[0, 'disj'])
    disj_cliente = f'DISJ. {disj}A'
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
    estrada()
    time.sleep(1)
    #cliente()
    #time.sleep(1)
    #prancha()
    print(f'Fim do desenho da ns {arquivo}')

print('Fim do Processo')
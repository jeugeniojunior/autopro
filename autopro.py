import pandas as pd
import pyautogui
import time
import openpyxl



def postes():
    print('Inserindo Postes')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('pcpb')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write('copiar')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write(coord_aglutinada_sem_inicial, 0.25)
    pyautogui.press('esc')

def rede():
    print('Inserindo Rede')
    pyautogui.moveTo(86, 283)
    pyautogui.click(86, 283)
    pyautogui.click(86, 283)
    pyautogui.write(coord_aglutinada)
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

def estrutura():
    print('Inserindo Estrutura')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('cm211300')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial_est)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write('copiar')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial_est)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(ponto_inicial_est)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write(coord_aglutinada_est, 0.25)
    pyautogui.press('esc')

def trafo():
    print('Inserindo Trafo')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('trafo')
    pyautogui.press('enter')
    pyautogui.write(ponto_trafo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(tipo_trafo)
    pyautogui.press('tab')
    pyautogui.press('enter')
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
    pyautogui.write(disj_cliente)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc')

def ns():
    print('Inserindo Numero de Ns')
    pyautogui.moveTo(545, 406)
    pyautogui.click(545, 406)
    pyautogui.write('dd')
    pyautogui.press('enter')
    pyautogui.write('t')
    pyautogui.press('enter')
    pyautogui.write(ponto_trafo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(numero_ns)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc')

#operação
print('AUTOPRO - AUTOMATIZADOR DE DESENHO DE PROJETO')

pyautogui.PAUSE = 1.5
pyautogui.alert('Não use o computador enquanto o programa funciona')

notas_df = pd.read_excel('notas.xlsx')

for i, notas in enumerate(notas_df['notas']):
    arquivo = notas_df.loc[i,'notas']
    planilha_df = pd.read_excel(f'{arquivo}.xlsx')
    coord_aglutinada = planilha_df.loc[0,'aglutinadas']
    coord_aglutinada_sem_inicial = planilha_df.loc[0,'aglutinadas menos inicial']
    coord_aglutinada_est = planilha_df.loc[0,'aglutinadas est']
    ponto_inicial = str(planilha_df.loc[0, 'coordenadas'])
    ponto_inicial_est = str(planilha_df.loc[0, 'estrutura'])
    ponto_trafo = str(planilha_df.loc[0, 'pto trafo'])
    tipo_trafo = str(planilha_df.loc[0, 'trafo'])
    nome_cliente = str(planilha_df.loc[0, 'cliente']).upper()
    telefone_cliente = str(planilha_df.loc[0, 'telefone'])
    disj = str(planilha_df.loc[0, 'disj'])
    disj_cliente = f'DISJ. {disj}A'
    numero_ns = str(planilha_df.loc[0, 'ns'])

    print(f'Inicio do {i+1} desenho')
    postes()
    time.sleep(2)
    rede()
    time.sleep(2)
    estrada()
    time.sleep(2)
    estrutura()
    time.sleep(2)
    trafo()
    time.sleep(2)
    cliente()
    time.sleep(2)
    ns()
    print(f'Fim do {i+1} desenho')

print('Fim do Processo')

pyautogui.alert('Fim do Processo')






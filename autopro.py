import pandas as pd
import pyautogui
import time
import openpyxl

planilha_df = pd.read_excel('coord_df.xlsx')
coord_aglutinada = planilha_df.loc[0,'aglutinadas']
coord_aglutinada_sem_inicial = planilha_df.loc[0,'aglutinadas menos inicial']
coord_aglutinada_est = planilha_df.loc[0,'aglutinadas est']
ponto_inicial = str(planilha_df.loc[0, 'coordenadas'])
ponto_inicial_est = str(planilha_df.loc[0, 'estrutura'])

def postes():
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
    pyautogui.moveTo(86, 283)
    pyautogui.click(86, 283)
    pyautogui.click(86, 283)
    pyautogui.write(coord_aglutinada)
    pyautogui.press('esc')

def estrada():
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

#operação

pyautogui.PAUSE = 1
pyautogui.alert('Não use o computador enquanto o programa fuciona')

postes()
time.sleep(2)
rede()
time.sleep(2)
estrada()
time.sleep(2)
estrutura()

import openpyxl
import pyperclip
import pyautogui
from time import sleep
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook['Produtos']
for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(927,239, duration=0.1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(921,325, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(917,418, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(933,483, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(942,555, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(909,625, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(905,673, duration=1)
    sleep(1)

    preço = linha[6].value
    pyperclip.copy(preço)
    pyautogui.click(972,263, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(964,329, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(978,396, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(951,467, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(947,546, duration=1)
    
    if tamanho == 'Pequeno':
        pyautogui.click(931,562, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(950,581, duration=1)
    else:
        pyautogui.click(949,601, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(973,603, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(906,652, duration=1)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(957,276, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(925,343, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(938,414, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigos_barra = linha[15].value
    pyperclip.copy(codigos_barra)
    pyautogui.click(899,521, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[15].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(930,590, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(906,640, duration=1)
    pyautogui.click(1245,198, duration=1)
    pyautogui.click(1245,198, duration=1)
    pyautogui.click(1075,459, duration=1)

# python -m venv envauto
# .\envauto\scripts\activate
# pip install pyautogui pandas openpyxl python-dotenv
# pip freeze > requirements.txt

import pyautogui, time, os, pandas as pd 
from dotenv import load_dotenv

# .env
load_dotenv()
email = os.getenv('EMAIL')
senha = os.getenv('SENHA')
link = os.getenv('LINK')

pyautogui.PAUSE = 0.5

# Abrir o navegador
pyautogui.press('win')
pyautogui.write('edge')
pyautogui.press('enter')
time.sleep(3)

# Abrir o site
pyautogui.write(link)
pyautogui.press('enter')
time.sleep(3)

# Fazer login
pyautogui.click(x=705, y=370)
pyautogui.write(email)
pyautogui.press('tab')
pyautogui.write(senha)
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(3)

df = pd.read_csv('produtos.csv')

# Cadastrar produtos
for i in df.index:

    # Código
    pyautogui.click(x=679, y=246)
    codigo = df.loc[i, 'codigo']
    pyautogui.write(str(codigo))

    # Marca
    pyautogui.press('tab')
    marca = df.loc[i, 'marca']
    pyautogui.write(str(marca))

    # Tipo
    pyautogui.press('tab')
    tipo = df.loc[i, 'tipo']
    pyautogui.write(str(tipo))

    # Categoria
    pyautogui.press('tab')
    categoria = df.loc[i, 'categoria']
    pyautogui.write(str(categoria))

    # Preço Unitário
    pyautogui.press('tab')
    preco_unitario = df.loc[i, 'preco_unitario']
    pyautogui.write(str(preco_unitario))

    # Custo
    pyautogui.press('tab')
    custo = df.loc[i, 'custo']
    pyautogui.write(str(custo))

    # Observações
    pyautogui.press('tab')
    obs = str(df.loc[i, 'obs'])
    if obs != "nan":
        pyautogui.write(obs)

    # Enviar
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('home')
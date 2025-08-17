import pandas as pd
import pyautogui
import time
import sys
import os
from pathlib import Path

# Configurações de segurança
pyautogui.FAILSAFE = True

# Configurações de tempo
DELAY_ENTRE_ACOES = 1
DELAY_PEQUENO = 0.3

def obter_coordenadas():
    """Retorna as coordenadas dos elementos na tela - ajuste conforme necessário"""
    return {
        'novo': (259, 286),
        'deposito': (283, 323),
        'data': (697, 421)
    }

def validar_dados(df):
    """Valida a estrutura do DataFrame"""
    colunas_necessarias = ['Data', 'Valor']
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(f"Coluna '{coluna}' não encontrada no Excel")

def processar_linha(data, valor):
    """Processa cada linha de dados"""
    print(f"\nInserindo: Data={data}, Valor={valor}")

    novo_pos = pyautogui.locateOnScreen('images/novo.png', confidence=0.9)
    if novo_pos:
        pyautogui.click(pyautogui.center(novo_pos))
        time.sleep(DELAY_PEQUENO)
        pyautogui.press('tab')
    else:
        print("Botão 'novo' não encontrado!")

    deposito_pos = pyautogui.locateOnScreen('images/deposito.png', confidence=0.9)
    if deposito_pos:
        pyautogui.click(pyautogui.center(deposito_pos))
        time.sleep(DELAY_PEQUENO)
        pyautogui.press('tab')
    else:
        print("Botão 'deposito' não encontrado!")

    # Segue com tab/digitação para o restante
    pyautogui.write(data)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(valor)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(DELAY_PEQUENO)
    # Verifica se valor termina com ',01' até ',10' e escreve o código correspondente (mapeamento personalizado)
    import re
    match = re.search(r',0([1-9]|10)$', valor)
    if match:
        sufixo = match.group(1)
        codigos = {
            '01': '0330',
            '02': '0333',
            '03': '1014',
            '04': '1454',
            '05': '0335',
            '06': '0336',
            '07': '0337',
            '08': '0338',
            '09': '0339',
            '10': '03310'
        }
        codigo = codigos.get(sufixo.zfill(2))
        if codigo:
            pyautogui.write(codigo)
            time.sleep(DELAY_PEQUENO)
            pyautogui.press('enter')

def encontrar_arquivo_downloads(nome_arquivo):
    """Encontra o arquivo na pasta Downloads do usuário"""
    # Obtém o caminho da pasta Downloads
    pasta_downloads = str(Path.home() / "Downloads")
    
    # Verifica se o arquivo existe com .xlsx ou .XLSX
    for arquivo in os.listdir(pasta_downloads):
        if arquivo.lower() == nome_arquivo.lower():
            return os.path.join(pasta_downloads, arquivo)
    
    raise FileNotFoundError(
        f"Arquivo '{nome_arquivo}' não encontrado na pasta Downloads. "
        f"Verifique se o arquivo existe e tem a extensão .xlsx"
    )

def main():
    try:
        # Verifica se o nome do arquivo foi fornecido
        if len(sys.argv) < 2:
            raise ValueError(
                "Uso: python bot.py nome_do_arquivo.xlsx\n"
                "Exemplo: python bot.py lancamentos.xlsx"
            )
        
        nome_arquivo = sys.argv[1]
        caminho_arquivo = encontrar_arquivo_downloads(nome_arquivo)
        
        # Configurações e preparação
        coords = obter_coordenadas()
        df = pd.read_excel(caminho_arquivo)
        validar_dados(df)
        
        print(f"\nBot iniciado - Arquivo: {nome_arquivo}")
        print(f"Total de registros: {len(df)}")
        print("Posicione o mouse na janela de destino em 5 segundos...")
        time.sleep(5)
        
        # Processamento dos dados
        for _, row in df.iterrows():
            data = row['Data'].strftime('%d/%m/%Y') if pd.notna(row['Data']) else ''
            valor_str = str(row['Valor']).replace(',', '.')
            valor = f"{float(valor_str):.2f}".replace(".", ",") if pd.notna(row['Valor']) else '0,00' 

            processar_linha(data, valor)
            time.sleep(DELAY_ENTRE_ACOES)
        
        print("\nProcessamento concluído com sucesso!")
    
    except Exception as e:
        print(f"\nERRO: {str(e)}")

if __name__ == "__main__":
    main()
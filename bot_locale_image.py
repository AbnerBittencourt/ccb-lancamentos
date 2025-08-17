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
TENTATIVAS_IMAGEM = 3
CONFIANCA_IMAGEM = 0.9  # Ajuste conforme necessário (0.8-0.95)

def encontrar_clicar(imagem, tentativas=TENTATIVAS_IMAGEM, confianca=CONFIANCA_IMAGEM):
    """Encontra e clica em um elemento na tela usando imagem"""
    for tentativa in range(tentativas):
        try:
            posicao = pyautogui.locateOnScreen(imagem, confidence=confianca)
            if posicao:
                centro = pyautogui.center(posicao)
                pyautogui.click(centro)
                return True
            time.sleep(0.5)
        except pyautogui.ImageNotFoundException:
            time.sleep(0.5)
    raise Exception(f"Elemento não encontrado: {imagem}")

def preencher_campo(imagem_campo, texto, limpar=True):
    """Preenche um campo de texto identificado por imagem"""
    if not encontrar_clicar(imagem_campo):
        return False
    
    if limpar:
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
    
    pyautogui.write(texto, interval=0.05)
    return True

def validar_dados(df):
    """Valida a estrutura do DataFrame"""
    colunas_necessarias = ['Data', 'Valor']
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(f"Coluna '{coluna}' não encontrada no Excel")

def processar_linha(data, valor):
    """Processa cada linha de dados"""
    print(f"\nInserindo: Data={data}, Valor={valor}")

    # Clica nos botões principais
    encontrar_clicar('imagens/novo.png')
    time.sleep(DELAY_PEQUENO)
    encontrar_clicar('imagens/deposito.png')
    time.sleep(DELAY_PEQUENO)

    # Preenche campos
    preencher_campo('imagens/campo_data.png', data)
    time.sleep(DELAY_PEQUENO)
    preencher_campo('imagens/campo_valor.png', valor)
    time.sleep(DELAY_PEQUENO)

    # Seleciona conta contábil
    encontrar_clicar('imagens/conta_contabil.png')
    time.sleep(DELAY_PEQUENO)
    encontrar_clicar('imagens/banco_brasil.png')
    time.sleep(DELAY_PEQUENO)

    # Condição especial para valores terminados em "01"
    if valor.endswith(",01"):
        print('Processando valor especial')
        encontrar_clicar('imagens/casa_oracao.png')
        time.sleep(DELAY_PEQUENO)
        pyautogui.write("0330")
        time.sleep(DELAY_PEQUENO)
        pyautogui.press('enter')

def encontrar_arquivo_downloads(nome_arquivo):
    """Encontra o arquivo na pasta Downloads do usuário"""
    pasta_downloads = str(Path.home() / "Downloads")
    
    for arquivo in os.listdir(pasta_downloads):
        if arquivo.lower() == nome_arquivo.lower():
            return os.path.join(pasta_downloads, arquivo)
    
    raise FileNotFoundError(
        f"Arquivo '{nome_arquivo}' não encontrado na pasta Downloads. "
        f"Verifique se o arquivo existe e tem a extensão .xlsx"
    )

def main():
    try:
        if len(sys.argv) < 2:
            raise ValueError(
                "Uso: python bot.py nome_do_arquivo.xlsx\n"
                "Exemplo: python bot.py lancamentos.xlsx"
            )
        
        nome_arquivo = sys.argv[1]
        caminho_arquivo = encontrar_arquivo_downloads(nome_arquivo)
        
        # Carrega e valida dados
        df = pd.read_excel(caminho_arquivo)
        validar_dados(df)
        
        print(f"\nBot iniciado - Arquivo: {nome_arquivo}")
        print(f"Total de registros: {len(df)}")
        print("Posicione a janela do sistema e aguarde 5 segundos...")
        time.sleep(5)
        
        # Processa cada linha
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
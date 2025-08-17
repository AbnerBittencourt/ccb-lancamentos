import pandas as pd
import pyautogui
import time
import sys
import os
import logging
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime

# Configurações de segurança
pyautogui.FAILSAFE = True

# Configurações de tempo
DELAY_ENTRE_ACOES = 1
DELAY_PEQUENO = 0.3
MAX_TENTATIVAS = 3
TIMEOUT_ELEMENTOS = 10

def configurar_logging():
    """Configura o sistema de logging"""
    log_dir = Path.home() / "bot_logs"
    log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / f"bot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

def selecionar_arquivo_interativamente():
    """Permite ao usuário selecionar o arquivo através de uma interface gráfica"""
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
        initialdir=str(Path.home() / "Downloads")
    )
    
    if not arquivo:
        raise ValueError("Nenhum arquivo selecionado")
    
    return arquivo

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

def esperar_elemento(imagem, timeout=TIMEOUT_ELEMENTOS, intervalo=0.5):
    """Espera até que um elemento apareça na tela"""
    inicio = time.time()
    while time.time() - inicio < timeout:
        pos = pyautogui.locateOnScreen(imagem, confidence=0.9)
        if pos:
            return pos
        time.sleep(intervalo)
    return None

def processar_linha(data, valor):
    """Processa cada linha de dados"""
    logging.info(f"Inserindo: Data={data}, Valor={valor}")

    # Botão 'novo'
    novo_pos = esperar_elemento('images/novo.png')
    if novo_pos:
        pyautogui.click(pyautogui.center(novo_pos))
        time.sleep(DELAY_PEQUENO)
        pyautogui.press('tab')
    else:
        raise Exception("Botão 'novo' não encontrado!")

    # Botão 'deposito'
    deposito_pos = esperar_elemento('images/deposito.png')
    if deposito_pos:
        pyautogui.click(pyautogui.center(deposito_pos))
        time.sleep(DELAY_PEQUENO)
        pyautogui.press('tab')
    else:
        raise Exception("Botão 'deposito' não encontrado!")

    # Preenchimento dos campos
    pyautogui.write(data)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(valor)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(DELAY_PEQUENO)
    
    # Verifica se valor termina com ',01' até ',10'
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

def processar_linha_com_tolerancia(data, valor, max_tentativas=MAX_TENTATIVAS):
    """Processa uma linha com múltiplas tentativas em caso de falha"""
    tentativa = 1
    while tentativa <= max_tentativas:
        try:
            processar_linha(data, valor)
            return True
        except Exception as e:
            logging.warning(f"Tentativa {tentativa} falhou: {str(e)}")
            tentativa += 1
            time.sleep(2)
    return False

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

def gerar_relatorio(sucessos, falhas):
    """Gera um relatório ao final da execução"""
    relatorio = {
        'total': sucessos + falhas,
        'sucessos': sucessos,
        'falhas': falhas,
        'taxa_sucesso': (sucessos / (sucessos + falhas)) * 100 if (sucessos + falhas) > 0 else 0
    }
    
    logging.info("\n===== RELATÓRIO DE EXECUÇÃO =====")
    logging.info(f"Total de registros: {relatorio['total']}")
    logging.info(f"Registros processados com sucesso: {relatorio['sucessos']}")
    logging.info(f"Registros com falha: {relatorio['falhas']}")
    logging.info(f"Taxa de sucesso: {relatorio['taxa_sucesso']:.2f}%")
    
    return relatorio

def main():
    configurar_logging()
    
    try:
        # Verifica se o nome do arquivo foi fornecido ou usa interface gráfica
        if len(sys.argv) < 2:
            caminho_arquivo = selecionar_arquivo_interativamente()
        else:
            nome_arquivo = sys.argv[1]
            caminho_arquivo = encontrar_arquivo_downloads(nome_arquivo)
        
        # Configurações e preparação
        coords = obter_coordenadas()
        df = pd.read_excel(caminho_arquivo)
        validar_dados(df)
        
        logging.info(f"\nBot iniciado - Arquivo: {caminho_arquivo}")
        logging.info(f"Total de registros: {len(df)}")
        logging.info("Posicione o mouse na janela de destino em 5 segundos...")
        time.sleep(5)
        
        # Processamento dos dados
        sucessos = 0
        falhas = 0
        
        for idx, row in df.iterrows():
            try:
                data = row['Data'].strftime('%d/%m/%Y') if pd.notna(row['Data']) else ''
                valor_str = str(row['Valor']).replace(',', '.')
                valor = f"{float(valor_str):.2f}".replace(".", ",") if pd.notna(row['Valor']) else '0,00' 

                if processar_linha_com_tolerancia(data, valor):
                    sucessos += 1
                else:
                    falhas += 1
                    logging.error(f"Falha ao processar linha {idx + 1}")
                
                time.sleep(DELAY_ENTRE_ACOES)
            
            except Exception as e:
                falhas += 1
                logging.error(f"Erro grave ao processar linha {idx + 1}: {str(e)}")
                time.sleep(2)  # Pausa maior em caso de erro
        
        relatorio = gerar_relatorio(sucessos, falhas)
        
        # Mostra popup com resumo (opcional)
        if relatorio['falhas'] > 0:
            messagebox.showwarning(
                "Processamento concluído com falhas",
                f"Sucessos: {relatorio['sucessos']}\nFalhas: {relatorio['falhas']}"
            )
        else:
            messagebox.showinfo(
                "Processamento concluído",
                f"Todos os {relatorio['sucessos']} registros foram processados com sucesso!"
            )
    
    except Exception as e:
        logging.error(f"\nERRO GRAVE: {str(e)}", exc_info=True)
        messagebox.showerror("Erro", f"Ocorreu um erro grave: {str(e)}")

if __name__ == "__main__":
    main()
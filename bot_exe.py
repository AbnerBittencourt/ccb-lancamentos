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

# ==============================================
# CONFIGURAÇÕES INICIAIS
# ==============================================

# Configura caminhos relativos
def obter_caminho_base():
    """Obtém o diretório base do aplicativo"""
    if getattr(sys, 'frozen', False):
        # Executando como .exe
        return Path(sys.executable).parent
    else:
        # Executando como script Python
        return Path(__file__).parent

BASE_DIR = obter_caminho_base()
IMAGES_DIR = BASE_DIR / "images"
LOGS_DIR = BASE_DIR / "bot_logs"

# Configurações de segurança
pyautogui.FAILSAFE = True

# Configurações de tempo
DELAY_ENTRE_ACOES = 1  # Delay entre ações (em segundos)
DELAY_PEQUENO = 0.3    # Delay pequeno para esperas curtas
MAX_TENTATIVAS = 3      # Máximo de tentativas por operação
TIMEOUT_ELEMENTOS = 10  # Tempo máximo para esperar um elemento

# ==============================================
# FUNÇÕES UTILITÁRIAS
# ==============================================

def centralizar_janela(janela):
    """Centraliza uma janela tkinter na tela"""
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f'+{x}+{y}')

def configurar_logging():
    """Configura o sistema de logging"""
    LOGS_DIR.mkdir(exist_ok=True)
    log_file = LOGS_DIR / f"bot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

def mostrar_mensagem(titulo, mensagem, tipo='info'):
    """Mostra uma mensagem centralizada na tela"""
    root = tk.Tk()
    root.withdraw()
    
    # Cria uma janela temporária para centralização
    temp_window = tk.Toplevel(root)
    temp_window.withdraw()
    
    if tipo == 'info':
        messagebox.showinfo(titulo, mensagem, parent=temp_window)
    elif tipo == 'warning':
        messagebox.showwarning(titulo, mensagem, parent=temp_window)
    else:
        messagebox.showerror(titulo, mensagem, parent=temp_window)
    
    temp_window.destroy()
    root.destroy()

def selecionar_arquivo_interativamente():
    """Permite ao usuário selecionar o arquivo através de uma interface gráfica"""
    root = tk.Tk()
    root.withdraw()
    
    # Configura a janela de diálogo
    janela_arquivo = tk.Toplevel()
    janela_arquivo.title("Selecione o arquivo Excel")
    centralizar_janela(janela_arquivo)
    janela_arquivo.withdraw()
    
    arquivo = filedialog.askopenfilename(
        parent=janela_arquivo,
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
        initialdir=str(Path.home() / "Downloads")
    )
    
    janela_arquivo.destroy()
    
    if not arquivo:
        raise ValueError("Nenhum arquivo selecionado")
    
    return arquivo

def encontrar_arquivo(nome_arquivo):
    """Encontra o arquivo na pasta Downloads ou no diretório atual"""
    # Verifica primeiro no diretório atual
    if os.path.exists(nome_arquivo):
        return nome_arquivo
    
    # Verifica na pasta Downloads
    pasta_downloads = str(Path.home() / "Downloads")
    for arquivo in os.listdir(pasta_downloads):
        if arquivo.lower() == nome_arquivo.lower():
            return os.path.join(pasta_downloads, arquivo)
    
    raise FileNotFoundError(
        f"Arquivo '{nome_arquivo}' não encontrado no diretório atual ou na pasta Downloads."
    )

# ==============================================
# FUNÇÕES PRINCIPAIS DO BOT
# ==============================================

def validar_dados(df):
    """Valida a estrutura do DataFrame"""
    colunas_necessarias = ['Data', 'Valor']
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(f"Coluna '{coluna}' não encontrada no Excel")

def esperar_elemento(imagem, timeout=TIMEOUT_ELEMENTOS, intervalo=0.5):
    """Espera até que um elemento apareça na tela"""
    caminho_imagem = str(IMAGES_DIR / imagem)
    inicio = time.time()
    
    while time.time() - inicio < timeout:
        pos = pyautogui.locateOnScreen(caminho_imagem, confidence=0.9)
        if pos:
            return pos
        time.sleep(intervalo)
    
    return None

def processar_linha(data, valor):
    """Processa cada linha de dados"""
    logging.info(f"Inserindo: Data={data}, Valor={valor}")

    # Botão 'novo'
    novo_pos = esperar_elemento("novo.png")
    if not novo_pos:
        raise Exception("Botão 'novo' não encontrado!")
    
    pyautogui.click(pyautogui.center(novo_pos))
    time.sleep(DELAY_PEQUENO)
    pyautogui.press('tab')

    # Botão 'deposito'
    deposito_pos = esperar_elemento("deposito.png")
    if not deposito_pos:
        raise Exception("Botão 'deposito' não encontrado!")
    
    time.sleep(DELAY_ENTRE_ACOES)
    pyautogui.click(pyautogui.center(deposito_pos))
    time.sleep(DELAY_PEQUENO)
    pyautogui.press('tab')

    # Preenchimento dos campos
    pyautogui.write(data)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(valor)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(DELAY_PEQUENO)
    
    # Verifica se valor termina com ',00' até ',10'
    match = re.search(r',0([0-9]|10)$', valor)
    if match:
        sufixo = match.group(1)
        codigos = {
            '00': '0330',
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

# ==============================================
# FUNÇÃO PRINCIPAL
# ==============================================

def main():
    try:
        configurar_logging()
        logging.info("Iniciando o bot...")
        
        # Verifica se foi passado um arquivo como argumento
        if len(sys.argv) > 1:
            nome_arquivo = sys.argv[1]
            caminho_arquivo = encontrar_arquivo(nome_arquivo)
        else:
            caminho_arquivo = selecionar_arquivo_interativamente()
        
        # Carrega e valida os dados
        df = pd.read_excel(caminho_arquivo)
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        validar_dados(df)
        
        logging.info(f"Arquivo carregado: {caminho_arquivo}")
        logging.info(f"Total de registros: {len(df)}")
        
        # Preparação para execução
        time.sleep(5)
        
        # Processamento dos dados
        sucessos = 0
        falhas = 0
        
        for idx, row in df.iterrows():
            try:
                # Formata os dados
                data = row['Data'].strftime('%d/%m/%Y') if pd.notna(row['Data']) else ''
                valor_str = str(row['Valor'])
                valor_str = re.sub(r'[^\d,.-]', '', valor_str).replace(',', '.')
                valor = f"{float(valor_str):.2f}".replace(".", ",") if pd.notna(row['Valor']) else '0,00'
                
                # Processa a linha
                if processar_linha_com_tolerancia(data, valor):
                    sucessos += 1
                else:
                    falhas += 1
                    logging.error(f"Falha ao processar linha {idx + 1}")
                
                time.sleep(DELAY_ENTRE_ACOES)
            
            except Exception as e:
                falhas += 1
                logging.error(f"Erro grave ao processar linha {idx + 1}: {str(e)}")
                time.sleep(2)
        
        # Gera e exibe o relatório
        relatorio = gerar_relatorio(sucessos, falhas)
        
        if relatorio['falhas'] > 0:
            mostrar_mensagem(
                "Processamento concluído com falhas",
                f"Sucessos: {relatorio['sucessos']}\nFalhas: {relatorio['falhas']}",
                "warning"
            )
        else:
            mostrar_mensagem(
                "Processamento concluído",
                f"Todos os {relatorio['sucessos']} registros foram processados com sucesso!",
                "info"
            )
    except Exception as e:
        logging.error(f"ERRO GRAVE: {str(e)}", exc_info=True)
        mostrar_mensagem("Erro", f"Ocorreu um erro grave: {str(e)}", "error")
        sys.exit(1)

if __name__ == "__main__":
    main()
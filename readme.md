# Passo a passo para rodar o bot de lançamentos do CCB

## 1. Instale o Python 3.12 ou superior

## 2. Clone o repositório

```sh
git clone https://github.com/AbnerBittencourt/ccb-lancamentos.git
cd ccb-lancamentos
```

## 3. Crie e ative o ambiente virtual

```sh
python3 -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate   # Windows
```

## 4. Instale as dependências

```sh
pip install -r requirements.txt
```

## 5. Instale dependências do sistema (Linux)

```sh
sudo apt-get update
sudo apt-get install python3-tk python3-dev gnome-screenshot
```

## 6. Execute o bot

```sh
python bot_coordinates.py nome_do_arquivo.xlsx
```

## 7. Ajuste as imagens e coordenadas conforme sua tela

# Passo a passo para executar no Windows

## 1. Instale o Python 3.12 ou superior

- Baixe em https://www.python.org/downloads/windows/
- Marque a opção "Add Python to PATH" na instalação.

## 2. Clone o repositório

```cmd
git clone https://github.com/AbnerBittencourt/ccb-lancamentos.git
cd ccb-lancamentos
```

## 3. Crie e ative o ambiente virtual

```cmd
python -m venv venv
venv\Scripts\activate
```

## 4. Instale as dependências

```cmd
pip install -r requirements.txt
```

## 5. Execute o bot

```cmd
python bot_coordinates.py nome_do_arquivo.xlsx
```

## 6. Observações

- O arquivo Excel deve estar na pasta Downloads do usuário.
- Ajuste as imagens e coordenadas conforme sua tela.
- Execute o script com a janela do sistema aberta e visível.

#

# Como gerar um executável (.exe) do bot no Windows

## 1. Instale o PyInstaller

```cmd
pip install pyinstaller
```

## 2. Gere o executável

```cmd
pyinstaller --onefile bot_exe.py
```

## 3. O arquivo gerado estará em `dist/bot_exe.exe`

- Basta executar clicando duas vezes ou pelo terminal:

```cmd
dist\bot_exe.exe nome_do_arquivo.xlsx
```

## 4. Observações

- O executável pode ser distribuído para outros computadores Windows sem precisar instalar Python.
- Certifique-se de copiar também as imagens da pasta `images/` se o bot depender delas.

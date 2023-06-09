#Bibliotecas
import os
import win32com.client as win32
import pythoncom
import shutil
from colorama import init, Fore
import time
import comtypes.client
from tqdm import tqdm


# Fecha todas as instâncias do Word antes de continuar
print(' ')
os.system("TASKKILL /F /IM winword.exe")
print(' ')
time.sleep(0.5)

#inicia a init do colorama
init()

#interação com usuario
print(' ')
print('Seja bem vindo ao programa de PDFs da ' + Fore.BLUE + 'IIndrA's')
print(' ')
time.sleep(0.5)
dir = input(Fore.WHITE + 'Insira o diretório: ')
print(' ')
#coleta o endereço dos arquivos
time.sleep(0.3)

print('você deseja inserir as pastas do setor de projetos cirion?')
print(' ')
past = input('digite S para SIM e N para NÃO: ')

if (past == 'S' or past == 's' or past == 'sim' or past =='SIM' or past == 'Sim'):

    # verifica se as pastas já existem
    dir_recebido = os.path.join(dir, "1 - Recebido Projetista")
    time.sleep(0.3)
    dir_scripts = os.path.join(dir, "2 - Scripts e Art")
    time.sleep(0.3)
    dir_envio = os.path.join(dir, "3 - Envio Gestor")
    time.sleep(0.3)
    dir_asbuilt = os.path.join(dir, "4 - AS-Built")
    time.sleep(0.3)
    if os.path.exists(dir_recebido) or os.path.exists(dir_scripts) or os.path.exists(dir_envio) or os.path.exists(dir_asbuilt):
        print(Fore.YELLOW + 'Atenção: uma ou mais pastas já existem neste diretório. Os arquivos podem ser sobrescritos.' + Fore.WHITE)
    print(' ')

    # cria as pastas no diretório principal se elas ainda não existem
    if not os.path.exists(dir_recebido):
        os.mkdir(dir_recebido)
    if not os.path.exists(dir_scripts):
        os.mkdir(dir_scripts)
    if not os.path.exists(dir_envio):
        os.mkdir(dir_envio)
    if not os.path.exists(dir_asbuilt):
        os.mkdir(dir_asbuilt)

else:
    print(' ')
    print('certo vamos continuar sem as pastas então')

print(' ')
print("Processando...")
print(' ')

# Defina o diretório de origem (onde os arquivos .docx e .doc estão)
dir_origem = dir

# Define e copia os arquivos para a pasta temporária
dir_temp = os.path.join(os.path.expanduser("~"), "Documents", "temparch")
if not os.path.exists(dir_temp):
    os.makedirs(dir_temp)

# Defina o diretório de destino (onde os arquivos .pdf serão salvos)
dir_destino = dir

# Inicialize a COM
pythoncom.CoInitialize()

# Cria uma instância do Word
word = win32.gencache.EnsureDispatch("Word.Application")

# Navegue pelo diretório de origem
for root, dirs, files in os.walk(dir_origem):
    for file in tqdm(files, desc='Convertendo arquivos'):
        if file.endswith(".doc") or file.endswith(".docx"):
            # Construa o caminho completo do arquivo de origem e destino
            arquivo_origem = os.path.join(root, file)
            arquivo_pdf = os.path.join(dir_destino, os.path.splitext(file)[0] + ".pdf")

            while True:
                try:
                    # Abra o arquivo e salve como PDF
                    document = word.Documents.Open(arquivo_origem)
                    document.SaveAs(arquivo_pdf, FileFormat=win32.constants.wdFormatPDF)
                    document.Close()
                    break  # sai do loop caso a conversão seja bem sucedida

                except Exception as e:
                    print(f"Erro ao processar o arquivo {arquivo_origem}: {e}")
                    time.sleep(1)  # espera 1 segundo antes de tentar novamente

# Feche o objeto do Word
word.Quit()
word = None

# Finalize a COM
pythoncom.CoUninitialize

# Move os arquivos gerados para a pasta de destino
for root, dirs, files in os.walk(dir_temp):
    for file in tqdm(files, desc='Movendo arquivos'):
        if file.endswith(".pdf"):
            arquivo_origem = os.path.join(root, file)
            arquivo_destino = os.path.join(dir_destino, file)
            shutil.move(arquivo_origem, arquivo_destino)

# Remove o diretório temporário
shutil.rmtree(dir_temp)

print(' ')
print('Todos os arquivos foram finalizados e convertidos, até a próxima!')
print(' ')
time.sleep(10)

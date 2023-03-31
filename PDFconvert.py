#Bibliotecas
import os
import shutil
from colorama import init, Fore
import time
from comtypes.client import Constants, CreateObject
from tqdm import tqdm

# Fecha todas as instâncias do Word antes de continuar
print(' ')
os.system("TASKKILL /F /IM winword.exe")
print(' ')
time.sleep(0.5)

#inicia a init do colorama
init()

print('Seja bem vindo ao conversor de' + Fore.RED + ' PDF')

#interação com usuario
print(' ')
print(' ')
time.sleep(0.5)
dir = input(Fore.WHITE + 'Insira o diretório: ')
print(' ')
#coleta o endereço dos arquivos
time.sleep(0.3)

print('você deseja inserir as pastas do setor de projetos cirion?')
print(' ')
past = ""

while True: 
    past = input(Fore.WHITE + 'digite ' + Fore.GREEN + ' S ' + Fore.WHITE + 'para SIM e' + Fore.YELLOW + ' N ' + Fore.WHITE + 'para NÃO: ')

    if past == "S" or past == "s":


        print("")
        print('Criando as pastas...')
        print("")

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
        break

    elif past == "N" or past == "n":
        print(' ')
        print('certo vamos continuar sem as pastas então')
        break

    else:
        
        print(" ")
        print('Opção' + Fore.RED + ' invalida' + Fore.WHITE + ', por favor escolha entre ' + Fore.YELLOW + 'N'+ Fore.WHITE + ' ou ' + Fore.GREEN + 'S')
        print(" ")

print(' ')
print(Fore.WHITE + "Processando...")
print(' ')

# Definir caminhos das pastas origem e temporaria
origem_path = dir
temporaria_path = "C:\\Users\\LeandroLourençoCorra\\Documents\\Temp"

# Criar pasta temporaria se ela não existir
if not os.path.exists(temporaria_path):
    os.makedirs(temporaria_path)

# Copiar arquivos .doc e .docx da pasta origem para a pasta temporaria
for filename in os.listdir(origem_path):
    if filename.endswith(('.doc', '.docx')):
        shutil.copy(os.path.join(origem_path, filename), temporaria_path)

# Converter arquivos .doc e .docx para .pdf usando o Microsoft Word
word = CreateObject('Word.Application')
word.Visible = False
for filename in tqdm(os.listdir(temporaria_path)):
    if filename.endswith(('.doc', '.docx')):
        doc_path = os.path.join(temporaria_path, filename)
        pdf_path = os.path.join(temporaria_path, os.path.splitext(filename)[0] + ".pdf")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, 17)
        doc.Close()
        
print(' ')
os.system("TASKKILL /F /IM winword.exe")
print(' ')
time.sleep(0.5)

# Mover arquivos .pdf da pasta temporaria para a pasta origem
for filename in os.listdir(temporaria_path):
    if filename.endswith('.pdf'):
        shutil.move(os.path.join(temporaria_path, filename), origem_path)

# Excluir pasta temporaria e seu conteúdo
shutil.rmtree(temporaria_path)

print(' ')
print('Todos os arquivos foram finalizados e convertidos, até a próxima!')
print(' ')
time.sleep(10)

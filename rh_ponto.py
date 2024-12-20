import os
import shutil
import time
import pyautogui
from datetime import datetime, timedelta
from send_email import enviar_email 

# Caminho do programa MD COM-LITE
program_path = r"C:\Program Files (x86)\Madis Rodbel\MD COM-LITE\MD COM-LITE.exe"

# Nome do programa a ser fechado
program_name = "MD COM-LITE.exe"

# Função para abrir o programa MD COM-LITE
def open_program(path):
    os.startfile(path)  # Abre o programa
    time.sleep(5)  # Aguarda carregamento completo
    print("MD COM-LITE iniciado com sucesso!")

# Função para fechar o programa MD COM-LITE
def close_program():
    os.system(f'taskkill /im "{program_name}" /f')  # Fecha o programa pelo nome
    print(f"{program_name} fechado com sucesso!")

# Função para conectar ao relógio ponto
def connect_to_time_clock(loop_count):
    time.sleep(2)  # Tempo para garantir que o programa esteja pronto

    # Acessar o menu de comunicação
    pyautogui.click(x=104, y=36)  # Clique no botão "Comunicação"
    print("Botão 'Comunicação' clicado!")

    # Navegar até a opção de conexão
    time.sleep(1)
    pyautogui.press("down", presses=2, interval=0.3)  # Navega para baixo
    pyautogui.press("right", presses=2, interval=0.3)  # Navega para a direita

    # Confirmar a conexão
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.click(x=636, y=379)  # Seleciona o relógio ponto
    if loop_count > 0:
        pyautogui.press("down", presses=loop_count, interval=0.3)  # Incrementa DOWN com base no loop
        print(f"DOWN pressionado {loop_count} vez(es).")
    pyautogui.press("enter", presses=3, interval=1.1)  # Confirma conexão
    print("Conexão com o relógio ponto confirmada!")

# Função para coletar marcações do relógio ponto
def coleta_marcacoes():
    time.sleep(2)  # Garantir que o programa esteja pronto

    # Acessar o menu de comunicação
    pyautogui.click(x=104, y=36)  # Clique no botão "Comunicação"
    print("Botão 'Comunicação' clicado!")

    # Navegar até a opção de coleta
    time.sleep(1)
    pyautogui.press("down", presses=2, interval=0.3)  # Navega para baixo
    pyautogui.press("right", presses=2, interval=0.3)  # Navega para a direita
    pyautogui.press("down", presses=2, interval=0.3)  # Mais navegações necessárias
    pyautogui.press("enter")  # Inicia a coleta

    # Finalizar coleta
    time.sleep(180)
    pyautogui.click(x=675, y=504)  # Clique no botão "Finalizar Coleta"
    print("Coleta de marcações finalizada!")

    # Desconectar do relógio
    time.sleep(2)
    pyautogui.click(x=104, y=36)  # Acessa novamente o menu
    pyautogui.press("down", presses=2, interval=0.3)  # Navega para baixo
    pyautogui.press("right", presses=2, interval=0.3)  # Navega para a direita
    pyautogui.press("enter")  # Confirma desconexão
    time.sleep(2)

# Função para calcular as datas
def calcular_datas():
    # Obter a data atual
    hoje = datetime.now()
    
    # Verificar se é segunda-feira
    '''
    if hoje.weekday() == 0:  # Segunda-feira
        data_inicio = hoje - timedelta(days=3)  # Sexta-feira anterior
        data_fim = hoje - timedelta(days=1)  # Domingo anterior
    else:
        data_inicio = hoje - timedelta(days=1)  # Dia anterior
        data_fim = hoje - timedelta(days=1)  # Dia anterior
    '''

    data_inicio = hoje - timedelta(days=1)  # Dia anterior
    data_fim = hoje - timedelta(days=1)  # Dia anterior

    # Formatar as datas no formato DD/MM/YYYY
    data_inicio_str = data_inicio.strftime("%d/%m/%Y")
    data_fim_str = data_fim.strftime("%d/%m/%Y")
    
    return data_inicio_str, data_fim_str

# Função para exportar marcações para arquivo TXT
def exporta_marcacoes():
    time.sleep(5)  # Espera o programa estar pronto
    pyautogui.click(x=391, y=61)  # Clique no botão Exportar
    time.sleep(1)

    file_names = ["pport2.txt", "pfab2.txt", "pfabNOVO.txt"]  # Lista de arquivos para exportação
    for file_name in file_names:
        pyautogui.press("enter")  # Confirma exportação inicial
        time.sleep(1)
        # Digitar o caminho completo para salvar o arquivo
        export_path = fr"\\172.25.*.*\system\ponto\{file_name}"
        pyautogui.write(export_path, interval=0.05)  # Escreve o caminho no campo
        print(f"Caminho exportado: {export_path}")
        time.sleep(0.3)
        pyautogui.press("enter")  # Confirma o nome do arquivo

        # Calcular as datas para exportação
        data_inicio, data_fim = calcular_datas()

        # Selecionar relógio para exportação
        pyautogui.press("tab")
        pyautogui.press("right")
        pyautogui.press("tab")
        pyautogui.press("enter")
        pyautogui.press("tab", presses=5, interval=0.1)
        pyautogui.write(data_inicio, interval=0.2)  # Inserir data inicial
        pyautogui.press("tab")
        pyautogui.write(data_fim, interval=0.2)  # Inserir data final
        pyautogui.press("tab")
        pyautogui.press("enter")
        time.sleep(60)
        pyautogui.press("left")
        pyautogui.press("enter", presses=2, interval=2)
        pyautogui.press("tab", presses=6, interval=0.2)
        pyautogui.press("right")
        pyautogui.press("tab", presses=8, interval=0.2)
        pyautogui.press("enter")
        pyautogui.press("tab", presses=8, interval=0.2)

# Função principal para executar o loop de conexão e coleta
def relogios_loop():
    for i in range(3):  # Repetir para 3 relógios diferentes
        print(f"Iniciando execução {i+1}/3 de conexão e coleta")
        connect_to_time_clock(loop_count=i)
        coleta_marcacoes()

def copiar_arquivos_txt(origem, destino):
    """
    Copia arquivos .txt da pasta de origem para a pasta de destino.

    :param origem: Caminho da pasta de origem.
    :param destino: Caminho da pasta de destino.
    """
    try:
        # Verifica se a pasta de destino existe, caso contrário cria.
        if not os.path.exists(destino):
            os.makedirs(destino)

        # Lista os arquivos na pasta de origem
        arquivos = os.listdir(origem)

        # Itera pelos arquivos e copia apenas os arquivos .txt
        for arquivo in arquivos:
            if arquivo.endswith('.txt'):
                caminho_origem = os.path.join(origem, arquivo)
                caminho_destino = os.path.join(destino, arquivo)
                shutil.copy2(caminho_origem, caminho_destino)
                print(f"Arquivo {arquivo} copiado para {destino}")
        print("Cópia concluída!")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    # Definir os caminhos de origem e destino
origem = r"\\172.25.*.*\system\ponto"
destino = r"\\172.25.*.*\system\ponto\bkp"

# Execução do programa
if __name__ == "__main__":
    copiar_arquivos_txt(origem, destino) # Faz bkp dos arquivos
    open_program(program_path)  # Abre o MD COM-LITE
    relogios_loop()  # Conecta e coleta informações
    exporta_marcacoes()  # Exporta as marcações
    close_program()  # Fecha o programa

   # Parâmetros do e-mail
    destinatario = "***.santos@***.ind.br, m****.***@***.ind.br, davidrosa@trajetoriaconsultoria.com.br"
    assunto = "RPA: Processo RH - Coleta e Exportação de Marcações de Ponto"
    mensagem_html = """
    <p>Prezado colaborador,</p>
    <p>Informo que o processo de <b>Coleta e Exportação de Marcações de Ponto</b> foi concluído com êxito pelo robô.</p>
    <p>Caso necessite de mais informações ou tenha alguma dúvida, entre em contato com a equipe de suporte da <u>Trajetória Consultoria.</u></p>

    <p>Atenciosamente,<br>RPA - *****</p>
    """
    enviar_email(destinatario, assunto, mensagem_html)  # Chamar a função de envio de e-mail

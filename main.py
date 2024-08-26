import sys
import tomllib

from rich.console import Console
from pathlib import Path

def libreoffice_checker():
    if Path("/usr/bin/soffice").exists():
        return Path("/usr/bin/soffice")
    return None

def config_file_loader():
    if (Path.cwd() / "config.ini").exists():
        with open(Path.cwd() / "config.ini", "rb") as _file:
            config = tomllib.load(_file)
            return config
    return None

def print_ascii():
    ascii_art = """
#    _____ _______ ______                                            
#   |  __ \__   __|  ____|                                           
#   | |__) | | |  | |__                                              
#   |  ___/  | |  |  __|    _____        _____   _____ ______ _____  
#   | | \ \  | |  | |      |  __ \_/\   |  __ \_/ ____|  ____|  __ \ 
#   |_|  \_\ |_|  |_|      | |__)_/  \  | |__) | (___ | |__  | |__) |
#                          |  ___/_/\_\ |  ___/ \___ \|  __| |  ___/ 
#                          | | _/ ____ \| | \ \ ____) | |____| | \ \ 
#   by Fabiano             |_|_/_/    \_\_|  \_\_____/|______|_|  \_\\
"""
    console = Console()
    console.print(ascii_art, style="bold red on black")
    print("Programa criado para auxiliar no processamento das máscaras de laudos, a sua\n"
          "utilização pode apresentar resultados inadequados que precisam ser validados\n"
          "antes de serem utilizados!\n")

def main_loop():
    console = Console()
    options = {"1": file_converter,
               "2": rtf_parser,
               "3": converter_and_parser,
               "4": program_exit}
    config = config_file_loader()
    try:
        wrong_option = 0
        while True:
            console.print("Escolha uma ação a ser realizada:", style="bold")
            console.print("[bold blue][ 1 ][/bold blue] Para realizar apenas a conversão de arquivos")
            console.print("[bold blue][ 2 ][/bold blue] Para realizar o parser dos arquivos RTF e remover as tags")
            console.print("[bold blue][ 3 ][/bold blue] Para converter e realizar o parser dos arquivos RTF")
            console.print("[bold blue][ 4 ][/bold blue] Para sair do programa")
            user_choice = console.input("\n[bold green]Sua opção é: ")
            if user_choice in options.keys():
                options[user_choice]()
                console.print("\nAção concluída\n", style="bold green")
            else:
                wrong_option += 1
                if wrong_option < 3:
                    console.print("\nOpção inválida, tente novamente\n", style="bold red")
                else:
                    print("Opa, parece que você não sabe usar este programa.\n"
                          "Pergunte a uma pessoa com mais conhecimento para te ajudar :]")
                    program_exit()
    except KeyboardInterrupt:
        program_exit()

def file_converter():
    print("somente conversão de arquivos")
    config = config_file_loader()
    if config:
        libreoffice_path = config["libreoffice"]["path"]
        print(libreoffice_path)
    else:
        libreoffice_path = ask_libreoffice_path()
        print(libreoffice_path)
    # todo: código da conversão de arquivos

def ask_libreoffice_path():
    console = Console()
    try:
        while True:
            console.print("Você precisa indicar o caminho da pasta 'program' do LibreOffice\n"
                          "Por exemplo: C:\\Program File\\LibreOffice\\program")
            user_path = console.input("Informe o caminho do LibreOffice: ")
            libre_path = Path(user_path) / "office.exe"
            if Path(libre_path).exists() and libre_path.is_file():
                return libre_path
            else:
                console.print("O caminho do LibreOffice informado não é válido ou não existe", style="bold red")
    except OSError:
        pass




def rtf_parser():
    print("parser do rtf")

def converter_and_parser():
    print("conversão e parser dos arquivos")

def program_exit():
    print("\nEncerrando o programa\n")
    sys.exit(0)

if __name__ == '__main__':
    print_ascii()
    main_loop()

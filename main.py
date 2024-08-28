import os
import subprocess
import sys
import time
import tomllib

from rich.console import Console
from rich.text import Text
from rich.live import Live
from rich.table import Table
from pathlib import Path

LIBRE_CMD_ARGS = ["--headless", "--convert-to", "rtf", "--outdir"]
EXTENSIONS = [".doc", ".docx", ".odt", ".html", ".txt"]
PATTERNS = [
    "{\\headerr",
    "{\\footerr",
    "{\\*\\themedata",
    "{\\*\\colorschememapping",
    "{\\*\\shppict",
    "{\\colortbl",
    "{\\stylesheet",
    "{\\*\\listtable",
    "{\\listoverridetable",
    "{\\*\\generator",
    "{\\info",
    "{\\*\\pgdsctbl",
    "{\\shp",
    "{\\nonshppict",
    "{\\pict",
    "{\\footer"
]

class RtfConversionError(Exception):
    pass

def print_ascii():
    ascii_art = """
#    _____ _______ ______                                            
#   |  __ \\__   __|  ____|                                           
#   | |__) | | |  | |__                                              
#   |  ___/  | |  |  __|    _____        _____   _____ ______ _____  
#   | | \\ \\  | |  | |      |  __ \\_/\\   |  __ \\_/ ____|  ____|  __ \\ 
#   |_|  \\_\\ |_|  |_|      | |__)_/  \\  | |__) | (___ | |__  | |__) |
#                          |  ___/_/\\_\\ |  ___/ \\___ \\|  __| |  ___/ 
#                          | | _/ ____ \\| | \\ \\ ____) | |____| | \\ \\ 
#   by Fabiano             |_|_/_/    \\_\\_|  \\_\\_____/|______|_|  \\_\\
"""
    console = Console()
    console.print(ascii_art, style="bold red on black")
    print("Programa criado para auxiliar no processamento das máscaras de laudos. A sua\n"
          "utilização pode apresentar resultados inadequados que precisam ser validados\n"
          "antes de serem utilizados!\n")

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


def loop_string(string: str, start):
    new_string = ""
    new_string += string[0:start]
    partial = parser(string[start:])
    new_string += partial
    return new_string


def parser(string: str):
    opened_brackets = 0
    closed_brackets = 0
    for i, c in enumerate(string):
        if c == "{":
            opened_brackets += 1
        if c == "}":
            closed_brackets += 1
        if opened_brackets == closed_brackets:
            return string[i+1:]
    return None

def extract_relative_folder(file_path: Path, root_folder:Path) -> str:
    parent = str(file_path.absolute().parent)
    relative = parent.replace(str(root_folder), "")
    if relative.startswith("\\"):
        relative = relative[1:]
    relative = relative.replace("\\", "-")
    if len(relative) < 1:
        relative = "Principal"
        # print(f"\n\n{relative}\n\n")
        # time.sleep(10)
    return relative


def main_loop():
    console = Console()
    options = {"1": file_converter,
               "2": rtf_parser,
               "3": converter_and_parser,
               "4": program_exit}
    config = config_file_loader()
    try:
        workdir = ask_for_work_dir()
        console.print(f"\nTrabalhando no diretório [bold white]{workdir}[/bold white]\n")

        wrong_option = 0
        while True:
            console.print("Escolha uma ação a ser realizada:", style="bold")
            console.print("[bold blue][ 1 ][/bold blue] Para realizar apenas a conversão de arquivos")
            console.print("[bold blue][ 2 ][/bold blue] Para realizar o parser dos arquivos RTF e remover as tags")
            console.print("[bold blue][ 3 ][/bold blue] Para converter e realizar o parser dos arquivos RTF")
            console.print("[bold blue][ 4 ][/bold blue] Gerar o arquivo CSV para importação das máscaras")
            console.print("[bold blue][ 5 ][/bold blue] Para sair do programa")
            user_choice = console.input("\n[bold green]Sua opção é: ")
            if user_choice in options.keys():
                message = options[user_choice](workdir)
                # console.print("\nAção concluída\n", style="bold green")
                console.print(message)
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

def ask_for_work_dir():
    console = Console()
    try:
        while True:
            console.print("Informe o diretório de trabalho para o programa, onde estão os arquivos "
                          "a serem convertidos ou parseados")
            user_path = console.input("[bold green]Informe o caminho do diretório: ")
            workdir = Path(user_path).absolute()
            if Path(workdir).exists() and workdir.is_dir():
                return workdir
            else:
                console.print("O caminho do diretório informado não é válido ou não existe\n", style="bold red")
    except OSError:
        pass

def file_converter(workdir: Path):
    console = Console()
    config = config_file_loader()
    if config:
        libreoffice_path = config["libreoffice"]["path"]
        conversion_args = config["libreoffice"]["args"]
        # print(libreoffice_path)
        # print(conversion_args)
    else:
        libreoffice_path = ask_libreoffice_path()
        conversion_args = LIBRE_CMD_ARGS
        # print(libreoffice_path)
        # print(conversion_args)
    soffice_cmd = [libreoffice_path] + conversion_args
    files_count = 0
    file_conv_ok = 0
    file_conv_error = 0
    console.print("\nPreparando conversão de arquivos para RTF ...\n")
    time.sleep(2)
    for item in workdir.glob(f"**/*"):
        if item.is_file() and item.stem.startswith("~"):
            pass
        if item.is_file() and item.suffix in config["extensions"]["ext"]:
            files_count += 1
            file_path = item.absolute()
            file_folder = item.parent
            try:
                command = soffice_cmd + [file_folder, file_path]
                proc_return = subprocess.run(command, capture_output=True, text=True)

                if proc_return.returncode == 0 and proc_return.stderr:
                    rel_path = extract_path_parts(item.absolute())
                    console.print(f"[bold red]ERRO[/bold red] ... conversão arquivo {item.suffix} => .rtf  => [magenta]{rel_path}")
                    file_conv_error += 1
                if proc_return.returncode == 0 and proc_return.stdout:
                    rel_path = extract_path_parts(item.absolute())
                    console.print(f"[bold green]SUCESSO[/bold green] ... conversão arquivo {item.suffix} => .rtf => [magenta]{rel_path}")
                    file_conv_ok += 1
            except:
                raise RtfConversionError("Erro inesperado na conversão de arquivos")
    message = Text()
    if file_conv_error > 0:
        message.append("\nConversão de arquivos finalizada com erros\n", style="bold red")
    else:
        message.append("\nConversão de arquivos finalizada com sucesso\n", style="bold green")
    message.append(f"Total de arquivos processados: ")
    message.append(f"{files_count}\n", style="bold blue")
    message.append(f"Total de arquivos convertidos: ")
    message.append(f"{file_conv_ok}\n", style="bold green")
    message.append(f"Total de arquivos não convertidos: ")
    message.append(f"{file_conv_error}\n", style="bold red")
    return message

def extract_path_parts(absolute_path: Path) -> str:
    parts = absolute_path.parts
    middle = int(len(parts) / 2)
    path = os.sep.join(parts[middle:])
    return "..." + os.sep + path


def ask_libreoffice_path():
    console = Console()
    try:
        while True:
            console.print("Você precisa indicar o caminho da pasta 'program' do LibreOffice\n"
                          "(por exemplo: C:\\Program File\\LibreOffice\\program)")
            user_path = console.input("[bold green]Informe o caminho do LibreOffice: [/bold green]")
            libre_path = Path(user_path) / "office.exe"
            if Path(libre_path).exists() and libre_path.is_file():
                return libre_path
            else:
                console.print("O caminho do LibreOffice informado não é válido ou não existe", style="bold red")
    except OSError:
        pass




def rtf_parser(workdir: Path):
    config = config_file_loader()
    table = Table(box=None)
    console = Console()

    files_count = 0
    console.print("\nPreparando parser dos arquivos RTF ...")
    time.sleep(2)
    with Live(table, refresh_per_second=4) as live:
        for item in workdir.glob(f"**/*.rtf"):
            with open(item, "r") as _file:
                rtf_content = _file.read()
                rtf_content = rtf_content.replace("\n", " ")

            for pattern in config["rtf"]["tags"]:
                while rtf_content.find(pattern) > 0:
                    pattern_index = rtf_content.find(pattern)
                    #  rtf_parsed = loop_string(rtf_content, pattern_index)
                    rtf_content = loop_string(rtf_content, pattern_index)

            with open(item, "w") as _file:
                _file.write(rtf_content)
                files_count += 1
                time.sleep(0.4)
                rel_path = extract_path_parts(item.absolute())
                table.add_row(f"Parser do arquivo [magenta]{rel_path} [bold green]concluído")
    message = Text()
    message.append("\nParser dos arquivos RTF concluído\n", style="bold green")
    message.append(f"Total de arquivos processados pelo parser: ")
    message.append(f"{files_count}\n", style="bold blue")
    return message

def converter_and_parser(workdir: Path):
    conv_message = file_converter(workdir)
    # time.sleep(3)
    parser_message = rtf_parser(workdir)
    message = Text()
    message.append(conv_message)
    message.append(parser_message)
    return message


def program_exit(arg=None):
    print("\n\nEncerrando o programa\n")
    sys.exit(0)

if __name__ == '__main__':
    print_ascii()
    main_loop()

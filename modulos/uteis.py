from colorama import Fore, Back, Style, init
from time import sleep


def cab(txt):
    init(autoreset=True)
    print(f'{Fore.WHITE + "="}' *50)
    print(f'{Fore.WHITE+Style.DIM+"|"}{Back.RED+Fore.WHITE+Style.BRIGHT+txt.center(48)}{Back.RESET+Fore.WHITE+Style.DIM+"|"}')
    print(f'{Fore.WHITE + "="}' *50)

def espace():
    print(' ')

def msg(txt):
    print(txt)

import os

def limpiar_pantalla():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')

def pausa():
    input("\nPresione Enter para continuar...")
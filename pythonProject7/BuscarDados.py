"""
    Modulo responsavel por leitura de planilhas e retorno dos dados no formato workbook
"""
# IMPORTANDO AS BIBLIOTECAS
from os import path
from openpyxl import load_workbook

# DECLARANDO AS FUNÇÕES
def obter_user_home():
    user_home = path.expanduser('~')
    return user_home


class BuscarDados:

    @staticmethod
    def ler_planilha(var):

        wb = load_workbook(F"{obter_user_home()}{var}")
        return wb



#!/usr/bin/python3

# Openpyxl
# pip3 install openpyxl
# https://openpyxl.readthedocs.io/en/stable/tutorial.html#playing-with-data

import os
from openpyxl import load_workbook

DATA_CSV = 'CorpusScoreRSE.csv'
SEP = '|'
DIR = 'Complet/'
file = 'Check-list - Diagnostic Région Occitanie - Pratiques environnementales et sociétales des entrep (1).xlsx'
SHEETNAME = 'liste des chapitres'
CELL_ENTREPRISE = 'A2'
CELL_SITE = 'A3'
# reférénces aux données dans les fichiers : col = colonne dans le fichier produit, li = ligne dans le fichier entrée
DATA_REFS = [
    {"col": "A1", "li": '10'},
    {"col": "A2", "li": '11'},
    {"col": "A3", "li": '12'},
    {"col": "A4", "li": '13'},
    {"col": "B5", "li": '15'},
    {"col": "B6", "li": '16'},
    {"col": "C7", "li": '18'},
    {"col": "C8", "li": '19'},
    {"col": "C9", "li": '20'},
    {"col": "C10", "li": '21'},
    {"col": "C11", "li": '22'},
    {"col": "C12", "li": '23'},
    {"col": "D13", "li": '25'},
    {"col": "D14", "li": '26'},
    {"col": "D15", "li": '27'},
    {"col": "E16", "li": '29'},
    {"col": "E17", "li": '30'},
    {"col": "F18", "li": '32'},
    {"col": "F19", "li": '33'},
    {"col": "F20", "li": '34'},
    {"col": "G21", "li": '36'},
    {"col": "G22", "li": '37'},
    {"col": "H23", "li": '39'},
    {"col": "H24", "li": '40'},
    {"col": "I25", "li": '42'},
    {"col": "I26", "li": '43'},
    {"col": "J27", "li": '45'},
    {"col": "J28", "li": '46'},
    {"col": "K29", "li": '48'},
    {"col": "K30", "li": '49'}
]

def get_num(fileEntry):
    tmp_nb = fileEntry.name.replace(').xlsx', '')
    nb = tmp_nb.split('(')[1]
    return int(nb)
    

with open(DATA_CSV,'w') as csv_file:
    header = 'Num' + SEP + 'Nom'
    for dr in DATA_REFS:
        col_name = dr["col"]
        header = header + SEP + col_name
    header = header + SEP + 'Lieu' + '\n'
    csv_file.write(header)

    with os.scandir(DIR) as entries:
        files = list(entries)
        files.sort(key=lambda x: get_num(x))
        for file in files:
            cur_file = DIR + file.name
            nb = get_num(file)
            wb = load_workbook(cur_file)
            sheet = wb[SHEETNAME]
            entreprise = sheet['A2'].value.replace('Entreprise:', '').strip()
            site = sheet['A3'].value.replace('Site:', '').replace('(', '').replace(')', '').replace('Siège','').strip()
            res_row = str(nb) + SEP + entreprise 
            for data_ref in DATA_REFS:
                cur_li = data_ref["li"]
                cur_cell = "D" + cur_li
                value = sheet[cur_cell].value
                if value == None:
                    data = ''
                else:
                    value = value.rstrip('%')
                    data = str(round(int(value) * 0.01, 2)).replace('.', ',')

                res_row = res_row + SEP + data
            res_row = res_row + SEP +  site + '\n'
            csv_file.write(res_row)

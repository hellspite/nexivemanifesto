import openpyxl

"""
Funzione per il caricamento del file excel e la selezione del foglio di calcolo
INPUT nome del file
OUTPUT False se il file è sconosciuto oppure restituisce il foglio di calcolo
"""


def load_excel(file_name):
    try:
        wb = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        print("File non trovato")
        return False

    return wb


"""
Funzione per selezionare il foglio di calcolo dal documento excel
"""


def select_sheet(wb):
    sheet_num = int(input("Seleziona il numero del foglio di calcolo: "))

    ws = None
    while ws is None:
        wb.active = (sheet_num - 1)
        if wb.active is None:
            print("Foglio di calcolo inesistente, provare di nuovo")

    return ws


"""
Funzione per creare un nuovo foglio excel che verrà poi caricato sul sito Nexive
"""


def create_empty_sheet():
    wb = openpyxl.Workbook()

    ws = wb.active

    ws['A1'] = "TITOLO"
    ws['B1'] = "COGNOME_RAGSOC"
    ws['C1'] = "NOME"
    ws['D1'] = "INFOUTILIXCONSEGNA"
    ws['E1'] = "VIA"
    ws['F1'] = "PIANOINTERNO"
    ws['G1'] = "CAP"
    ws['H1'] = "FRAZIONE"
    ws['I1'] = "COMUNE"
    ws['J1'] = "PROVINCIA"
    ws['K1'] = "IMPORTO_CONTRASSEGNO"
    ws['L1'] = "MODALITACONTRASSEGNO"
    ws['M1'] = "COLLI"
    ws['N1'] = "TAGLIA"
    ws['O1'] = "CELLULARE"
    ws['P1'] = "TELEFONO"
    ws['Q1'] = "EMAIL"
    ws['R1'] = "RIFERIMENTO_CLIENTE"
    ws['S1'] = "CENTRO_COSTO"
    ws['T1'] = "MITTENTE"
    ws['U1'] = "MITTENTE_NOME"
    ws['V1'] = "MITTENTE_VIA"
    ws['W1'] = "MITTENTE_CAP"
    ws['X1'] = "MITTENTE_COMUNE"
    ws['Y1'] = "MITTENTE_PROVINCIA"
    ws['Z1'] = "IMPORTOASSICURATA"
    ws['AA1'] = "PRODOTTO"
    ws['AB1'] = "SERVIZIO_RESI_EASY"

    return wb


"""
Funzione per contare le righe effettive da processare
"""


def count_rows(worksheet):
    col_a = worksheet['A']

    row_num: int = 0

    for i in col_a:
        if i.value is not None:
            row_num += 1

    return row_num


"""
Funzione che si occupa dell'elaborazione del contenuto
"""


def parse_content(quantity, content):

    content = content.lower()
    content = content.replace('maglietta io rompo black', 'B')
    content = content.replace('maglietta io rompo orange', 'O')
    content = content.replace('femmina', 'F')
    content = content.replace('maschio', 'M')
    content = content.upper()

    if quantity > 1:
        content = content + "   "
        content = content * quantity
        content = content.rstrip()

    return content


"""
Funzione per passare i dati dal primo foglio excel a quello definitivo per la spedizione Nexive
"""


def parse_xl(ws_in, wb_out):

    file_in = ws_in
    file_out = wb_out.active

    rows = count_rows(file_in)

    # Numero d'ordine
    for i in range(3, rows):
        file_out['A' + str(i - 1)] = file_in['A' + str(i)]

    # Email
    for i in range(3, rows):
        file_out['Q' + str(i - 1)] = file_in['D' + str(i)]

    # Nome
    for i in range(3, rows):
        file_out['B' + str(i - 1)] = file_in['E' + str(i)]

    # Indirizzo
    for i in range(3, rows):
        file_out['E' + str(i - 1)] = file_in['F' + str(i)]

    # Presso
    for i in range(3, rows):
        file_out['C' + str(i - 1)] = file_in['G' + str(i)]

    # Città
    for i in range(3, rows):
        file_out['I' + str(i - 1)] = file_in['H' + str(i)]

    # CAP
    for i in range(3, rows):
        cap = str(file_in['I' + str(i)])
        cap_len = len(cap)
        if cap_len < 5:
            for c in range(cap_len, 5):
                cap = '0'+cap

        file_out['G' + str(i - 1)] = cap

    # Provincia
    for i in range(3, rows):
        file_out['J' + str(i - 1)] = file_in['J' + str(i)]

    # Telefono
    for i in range(3, rows):
        file_out['P' + str(i - 1)] = file_in['K' + str(i)]

    # Taglia
    for i in range(2, rows):
        file_out['N' + str(i - 1)] = 'S'

    # Contenuto
    for i in range(3, rows):
        quantity = file_in['B' + str(i)]
        shirt = file_in['C' + str(i)]

        shirt = parse_content(quantity, shirt)

        file_out['D' + str(i - 1)] = shirt

    return wb_out


"""
Funzione per salvare il file excel finale
"""

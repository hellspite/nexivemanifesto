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

    sheet_num = int(input("Seleziona il numero del foglio di calcolo: "))

    ws = None
    while ws is None:
        wb.active = (sheet_num - 1)
        if wb.active is None:
            print("Foglio di calcolo inesistente, provare di nuovo")

        ws = wb.active

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

    return ws


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
Funzione per passare i dati dal primo foglio excel a quello definitivo per la spedizione Nexive
"""


def parse_xml(in_xl):
    # Carico il file in entrata
    file_in = load_excel(in_xl)

    # Creo il file in uscita
    file_out = create_empty_sheet()

    rows = count_rows(file_in)

    # Numero d'ordine
    for i in range(3, rows):
        file_out['A' + str(i - 1)] = file_in['A' + str(i)]

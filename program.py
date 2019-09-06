import parsexl

"""
Il programma principale
"""

if __name__ == "__main__":
    while True:
        file_name = None
        while file_name is None:
            try:
                file_name = input('Nome del file excel: ')
            except FileNotFoundError:
                print("File non trovato, riprovare")

        wb_in = parsexl.load_excel(file_name)

        wb_out = parsexl.create_empty_sheet()

        ws_in = parsexl.select_sheet(wb_in)

        wb_nexive = parsexl.parse_xl(ws_in, wb_out)

        wb_nexive.save('nexive.xlsx')

        break

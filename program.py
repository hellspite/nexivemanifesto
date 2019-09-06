import parsexl

"""
Il programma principale
"""

if __name__ == "__main__":
    while True:
        ok_go = False
        while ok_go is False:
            file_name = input('Nome del file excel: ')

            try:
                wb_in = parsexl.load_excel(file_name)
            except:
                print("File non valido, riprovare")
            else:
                if wb_in is not False:
                    ok_go = True

        wb_out = parsexl.create_empty_sheet()

        ws_in = parsexl.select_sheet(wb_in)

        wb_nexive = parsexl.parse_xl(ws_in, wb_out)

        wb_nexive = parsexl.check_rows(wb_nexive)

        wb_nexive.save('nexive.xlsx')

        break

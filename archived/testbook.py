

import xlwings as xw

app = xw.App(visible=False)
wb = app.books.open('../test_book.xlsx')
ws = wb.sheets[0]

for n in range(1, 20):
    print(ws.range(f'B{n}').value)

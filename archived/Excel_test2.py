
import xlwings as xw

app = xw.App(visible=False)
wb = app.books.open('mkup_excel.xlsx')

num_sheets = len(wb.sheets)
if not num_sheets == 1:
    for n in range(num_sheets-1, 0, -1):
        wb.sheets[n].delete()

orig = wb.sheets[0]

mdh = wb.sheets.add(name="MDH")
tin = wb.sheets.add(name="TIN")
like = wb.sheets.add(name="LIKE")
alma = wb.sheets.add(name="ALMA")

colw = 24
rowh = 116

#print(orig.range('C4').row_height)
mdh.range('C:G').column_width = colw
tin.range('C:G').column_width = colw
like.range('C:G').column_width = colw
alma.range('C:G').column_width = colw

def SetColNames(sheet, orig):
    sheet.range('A1').value = orig.range('A3').value
    sheet.range('B1:F1').value = orig.range('C3:G3').value
    sheet.range('G1:I1').value = orig.range('U3:X3').value
    sheet.range('J1').value = orig.range('AD3').value
    sheet.range('K1:L1').value = orig.range('AI3:AJ3').value
    sheet.range('L1').value = 'Costo'
    sheet.range('M1').value = 'sugerido'


SetColNames(sheet=mdh, orig=orig)
SetColNames(sheet=tin, orig=orig)
SetColNames(sheet=like, orig=orig)
SetColNames(sheet=alma, orig=orig)


def CopyRow(origin, destino, row_origin, row_destino):
    destino.range(f'A{row_destino}').row_height = 116
    destino.range(f'A{row_destino}').value = origin.range(f'A{row_origin}').value
    destino.range(f'B{row_destino}:F{row_destino}').value = origin.range(f'C{row_origin}:{row_origin}').value
    destino.range(f'G{row_destino}:I{row_destino}').value = origin.range(f'{row_origin}:X{row_origin}').value
    destino.range(f'{row_destino}').value = origin.range(f'{row_origin}').value
    destino.range(f'{row_destino}:L{row_destino}').value = origin.range(f'{row_origin}:{row_origin}').value


finished = False
r = 4  # row count in origin
rmdh = 2
rtin = 2
rlike = 2
ralma = 2

while not finished:
    bodega = orig.range(f'A{r}').value

    if bodega == "":
        pass

    match bodega:
        case "MDH":
            CopyRow(origin=orig, destino=mdh,row_origin=r, row_destino=rmdh)
            rmdh = rmdh + 1

        case "TIN":
            CopyRow(origin=orig, destino=tin, row_origin=r, row_destino=rtin)
            rtin = rtin + 1

        case "LIKE":
            CopyRow(origin=orig, destino=like, row_origin=r, row_destino=rlike)
            rlike = rlike + 1

        case "ALMA BEAUTY":
            CopyRow(origin=orig, destino=alma, row_origin=r, row_destino=ralma)
            ralma = ralma + 1


    r = r + 1


#wb.save('mod.xlsx')
wb.close()
app.quit()


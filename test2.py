
import xlwings as xw

app = xw.App(visible=False)
wb = app.books.open('mkup_excel.xlsx')

ws1 = wb.sheets[0]

ws2 = wb.sheets.add(name="MDH")


colw = 24
rowh = 116

#print(orig.range('C4').row_height)
ws2.range('B:F').column_width = colw



def set_col_names(sheet, origin):
    sheet.range('A1').value = origin.range('A3').value
    sheet.range('B1:F1').value = origin.range('C3:G3').value
    sheet.range('G1:J1').value = origin.range('U3:X3').value

    sheet.range('H1:I1').api.Merge()
    sheet.range('H1:I1').api.HorizontalAlignment = -4108

    sheet.range('K1').value = origin.range('AD3').value
    sheet.range('L1:M1').value = origin.range('AI3:AJ3').value
    sheet.range('N1').value = 'Costo'
    sheet.range('O1').value = 'Sugerido'


set_col_names(ws2, ws1)


ws2.range('A2').value = ws1.range('A4').value
#ws2.range('B2:F2').value = ws1.range('C4:G4').value

print(ws1.pictures[0].top_left_cell.address)

#wb.save('mod.xlsx')


wb.close()
app.quit()


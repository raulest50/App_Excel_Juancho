

import xlwings as xw

def CopyRow(value, work_sheet, row_number):
    match value:
        case "mdh":
            work_sheet.range(f'C{row_number}').value = value
        case "tin":
            work_sheet.range(f'D{row_number}').value = value
        case "like":
            work_sheet.range(f'E{row_number}').value = value
        case "alma":
            work_sheet.range(f'F{row_number}').value = value


def CopyRow2(value, work_sheet, row_number):
    match value:
        case "mdh":
            work_sheet.range(f'E{row_number}:F{row_number}').value = work_sheet.range(f'B{row_number}:C{row_number}').value
        case "tin":
            work_sheet.range(f'E{row_number}:F{row_number}').value = work_sheet.range(f'B{row_number}:C{row_number}').value
        case "like":
            work_sheet.range(f'E{row_number}:F{row_number}').value = work_sheet.range(f'B{row_number}:C{row_number}').value
        case "alma":
            work_sheet.range(f'E{row_number}:F{row_number}').value = work_sheet.range(f'B{row_number}:C{row_number}').value


app = xw.App(visible=True)
wb = app.books.open('../test_book2.xlsx')
ws = wb.sheets[0]

n = 2

finished = False
while not finished:
    c = ws.range(f'B{n}')
    v = c.value
    if v is None:
        finished = True

    elif c.api.MergeCells:  # if merged cells
        ma = c.merge_area.value
        nma = len(ma)
        v = [it for it in ma if it is not None][0]
        for k in range(n, n + nma):
            CopyRow2(value=v, work_sheet=ws, row_number=k)
        range_2_merge = ws.range(f'E{n}:E{n + nma-1}')
        range_2_merge.api.Merge()
        range_2_merge.api.HorizontalAlignment = -4108  # -4108 is the code for center alignment
        n = n + nma

    elif not c.api.MergeCells:
        CopyRow2(value=v, work_sheet=ws, row_number=n)
        n = n+1


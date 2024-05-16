
import xlwings as xw


class ExcelHandler:

    col_width = 24
    row_height = 116

    cel_marca_orig = 'A3'
    cel_fotos_orig = 'C3:G3'  # son 5 celdas con fotos
    cells_ccp_orig = 'U3:X3'  # ctns, cantidad y precio RMB und

    cols_cantidad_destino = 'V1:W1' #   solo para hacer merge en los sheets de destino

    cel_total_orig = 'AD3'
    cel_ct_orig = 'AI3:AJ3'


    def __init__(self, file_path, file_destino):

        # index counter para cada una de las hojas destino y la fuente
        start_dest_row = 1
        self.n_mdh = start_dest_row
        self.n_tin = start_dest_row
        self.n_like = start_dest_row
        self.n_alma = start_dest_row
        self.n_orig = 4

        self.file_path = file_path
        self.file_destino = file_destino
        self.app = xw.App(visible=False)
        self.wb = self.app.books.open(file_path)
        self.orig = self.wb.sheets[0]

        num_sheets = len(self.wb.sheets)
        if not num_sheets == 1:
            for n in range(num_sheets - 1, 0, -1):
                self.wb.sheets[n].delete()

        # Initialize sheets as attributes
        self.mdh = self.wb.sheets.add(name="MDH")
        self.tin = self.wb.sheets.add(name="TIN")
        self.like = self.wb.sheets.add(name="LIKE")
        self.alma = self.wb.sheets.add(name="ALMA BEAUTY")

        self.setup_sheets()


    # para cada hoja se ajustan anchos de columna y se ponen los titulos
    def setup_sheets(self):
        for sheet in [self.mdh, self.tin, self.like, self.alma]:
            sheet.range('B:F').column_width = self.col_width
            self.set_col_names(sheet)


    def set_col_names(self, sheet):
        sheet.range('A1').value = self.orig.range(self.cel_marca_orig).value
        sheet.range('B1:F1').value = self.orig.range(self.cel_fotos_orig).value
        sheet.range('G1:J1').value = self.orig.range(self.cells_ccp_orig).value
        self.merge_center(sheet.range('H1:I1'))

        sheet.range('K1').value = self.orig.range('AD3').value
        sheet.range('L1:M1').value = self.orig.range('AI3:AJ3').value
        sheet.range('N1').value = 'Costo'
        sheet.range('O1').value = 'Sugerido'

    def save_and_close(self):
        self.wb.save(self.file_destino)
        self.wb.close()
        self.app.quit()

    def merge_center(self, range):
        range.api.Merge()
        range.api.HorizontalAlignment = -4108  # -4108 is the code for center alignment

    def CopyRow(self, sheet, row_destino, row_origin):
        sheet.range(f'A{row_destino}').row_height = self.row_height
        sheet.range(f'A{row_destino}').value = self.orig.range(f'A{row_origin}').value
        sheet.range(f'B{row_destino}:F{row_destino}').value = self.orig.range(f'C{row_origin}:{row_origin}').value
        sheet.range(f'G{row_destino}:I{row_destino}').value = self.orig.range(f'{row_origin}:X{row_origin}').value
        sheet.range(f'{row_destino}').value = self.orig.range(f'{row_origin}').value
        sheet.range(f'{row_destino}:L{row_destino}').value = self.orig.range(f'{row_origin}:{row_origin}').value


    def DoRow(self, value, row_origin):
        match value:
            case "mdh":
                self.CopyRow(self.mdh, row_destino=self.n_mdh, row_origin=row_origin)
                self.n_mdh = self.n_mdh + 1 # siempre se avanza 1
            case "tin":
                self.CopyRow(self.tin, row_destino=self.n_tin, row_origin=row_origin)
                self.n_tin = self.n_tin + 1 # si se invoca dede el elif del merge
            case "like":
                self.CopyRow(self.like, row_destino=self.n_like, row_origin=row_origin)
                self.n_like = self.n_like + 1 # se incrementa nma veces +1
            case "alma":
                self.CopyRow(self.alma, row_destino=self.n_alma, row_origin=row_origin)
                self.n_alma = self.n_alma + 1


    def ProcesarWorkBook(self):
        finished = False
        while not finished:
            c = self.orig.range(f'B{self.n_orig}')  #  celda que indica bodega corresponde el row
            v = c.value
            if v is None:  # Si llega a una celda vacia llego al fin de los datos
                finished = True

            elif c.api.MergeCells:  # if merged cells
                ma = c.merge_area.value
                nma = len(ma)
                v = [it for it in ma if it is not None][0]
                for k in range(self.n_orig, self.n_orig + nma):
                    self.DoRow(value=v, row_origin=k)
                to = self.n_orig + nma - 1
                self.MergeBodegaColumn(value=v, fro=self.n_orig, to=to)
                self.n_orig = self.n_orig + nma  # se avanza por todas las celdas que estan en el merge

            elif not c.api.MergeCells: # si es una sola celda
                self.DoRow(value=v, row_origin=self.n_orig)
                self.n_orig = self.n_orig + 1  # se avanza solo una celda
        self.save_and_close()


    def MergeBodegaColumn(self, value, fro, to):
        match value:
            case "mdh":
                range_2_merge = self.mdh.range(f'E{fro}:E{to}')
                self.merge_center(range=range_2_merge)
            case "tin":
                range_2_merge = self.tin.range(f'E{fro}:E{to}')
                self.merge_center(range=range_2_merge)
            case "like":
                range_2_merge = self.like.range(f'E{fro}:E{to}')
                self.merge_center(range=range_2_merge)
            case "alma":
                range_2_merge = self.alma.range(f'E{fro}:E{to}')
                self.merge_center(range=range_2_merge)



eh = ExcelHandler('mkup_excel.xlsx', 'mod.xlsx')
#eh.mdh.range('A2:A4').value = 10
#print(eh.orig.range('H4:H5').value)
eh.save_and_close()




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


    # para cada hoja se ajustan anchos de columna y se ponen los titulos
    def setup_sheets(self):
        for sheet in [self.mdh, self.tin, self.like, self.alma]:
            sheet.range('C:G').column_width = self.col_width
            self.set_col_names(sheet)
            sheet.range()


    def set_col_names(self, sheet):
        sheet.range('A1').value = self.orig.range(self.cel_marca_orig).value
        sheet.range('B1:F1').value = self.orig.range(self.cel_fotos_orig).value
        sheet.range('G1:I1').value = self.orig.range(self.cells_ccp_orig).value

        self.merge_center(sheet.range(self.cols_cantidad_destino)) ## se jutan V y W en el titulo solamente

        sheet.range('J1').value = self.orig.range('AD3').value
        sheet.range('K1:L1').value = self.orig.range('AI3:AJ3').value
        sheet.range('L1').value = 'Costo'
        sheet.range('M1').value = 'Sugerido'

    def save_and_close(self):
        self.wb.save(self.file_destino)
        self.wb.close()
        self.app.quit()

    def merge_center(self, range):
        range.api.Merge()
        range.api.HorizontalAlignment = -4108  # -4108 is the code for center alignment

    def CopyRow(self, sheet, row_destino, row_origin):
        sheet.range(f'A{row_destino}').row_height = 116
        sheet.range(f'A{row_destino}').value = self.orig.range(f'A{row_origin}').value
        sheet.range(f'B{row_destino}:F{row_destino}').value = self.orig.range(f'C{row_origin}:{row_origin}').value
        sheet.range(f'G{row_destino}:I{row_destino}').value = self.orig.range(f'{row_origin}:X{row_origin}').value
        sheet.range(f'{row_destino}').value = self.orig.range(f'{row_origin}').value
        sheet.range(f'{row_destino}:L{row_destino}').value = self.orig.range(f'{row_origin}:{row_origin}').value
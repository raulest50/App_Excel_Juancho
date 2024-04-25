Attribute VB_Name = "CustomSubs"

Function CreateSheetIfNotExist(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False

    ' Check if the sheet already exists
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If the sheet does not exist, create it
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    
    ' Return the worksheet
    Set CreateSheetIfNotExist = ws
End Function


Function CountMergedCells(ws As Worksheet, row As Integer, column As Integer) As Integer
    Dim cell As Range
    Set cell = ws.Cells(row, column)

    ' Check if the cell is part of a merged range
    If cell.MergeCells Then
        ' Return the total number of cells in the merged range
        CountMergedCells = cell.MergeArea.Count
    Else
        ' Return 1 if the cell is not merged
        CountMergedCells = 1
    End If
End Function


Sub CleanWorkBook()
    Dim ws As Worksheet
    Dim safeSheets As Collection
    Set safeSheets = New Collection

    ' List of sheet names to keep
    safeSheets.Add "Despacho"
    safeSheets.Add "MDH"
    safeSheets.Add "TIN"
    safeSheets.Add "LIKE"
    safeSheets.Add "ALMA"

    ' Disable alerts to prevent Excel from asking confirmation for each sheet delete
    Application.DisplayAlerts = False
    
    ' Loop through each sheet in the workbook in reverse order to avoid skipping sheets after deletion
    Dim i As Integer
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        If Not SheetExistsInCollection(ws.Name, safeSheets) Then
            ws.Delete
        End If
    Next i
    
    ' Re-enable alerts after operations are done
    Application.DisplayAlerts = True

    MsgBox "Sheets cleanup complete."
End Sub


Sub DeleteAllExceptDespacho()
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Disable warning messages

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Despacho" Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True ' Re-enable warning messages
End Sub


Function SheetExistsInCollection(sheetName As String, coll As Collection) As Boolean
    Dim itm
    On Error Resume Next ' Ignore errors which happen when an item isn't found
    For Each itm In coll
        If itm = sheetName Then
            SheetExistsInCollection = True
            Exit Function
        End If
    Next itm
    SheetExistsInCollection = False ' Return false if not found
    On Error GoTo 0 ' Turn off error ignoring
End Function



' this is used only to know the col width and row heigth numbers
' but it is not intended to be part of the actual algorithm, it is more like both
' GetCellDimensions and ColumnLetter are helper routines for getting the numbers
' to put in the actual algorithms
Sub GetCellDimensions(ws As Worksheet, inputRow As Integer, inputColumn As Integer)
    ' Validate if the worksheet object is set
    If ws Is Nothing Then
        MsgBox "Invalid worksheet reference provided.", vbCritical
        Exit Sub
    End If
    
    ' Retrieve column width in character units and points
    Dim columnWidthCharacters As Double
    columnWidthCharacters = ws.Columns(inputColumn).ColumnWidth
    
    Dim columnWidthPoints As Double
    columnWidthPoints = ws.Columns(inputColumn).Width
    
    ' Retrieve row height in points
    Dim rowHeight As Double
    rowHeight = ws.Rows(inputRow).Height
    
    ' Displaying the results in a message box
    MsgBox "Dimensions for Column " & ColumnLetter(inputColumn) & " (" & inputColumn & ") and Row " & inputRow & ":" & vbCrLf & _
           "Column Width: " & columnWidthCharacters & " characters (" & columnWidthPoints & " points)" & vbCrLf & _
           "Row Height: " & rowHeight & " points", vbInformation, "Cell Dimensions in " & ws.Name
End Sub

' Helper function to convert column number to letter (optional for enhanced readability)
Function ColumnLetter(colNum As Integer) As String
    Dim v As Integer
    Dim s As String
    v = colNum
    s = ""
    While v > 0
        Dim m As Integer
        m = (v - 1) Mod 26
        s = Chr(65 + m) & s
        v = Int((v - m) / 26)
    Wend
    ColumnLetter = s
End Function




Sub SetupColTitles(hoja_main As Worksheet, otherSheet As Worksheet)
    ' Validate if the worksheet objects are set
    If hoja_main Is Nothing Or otherSheet Is Nothing Then
        MsgBox "One or more invalid worksheet references provided.", vbCritical
        Exit Sub
    End If
    
    otherSheet.Rows(1).rowHeight = 59

    ' Set the width of columns in otherSheet
    otherSheet.Columns(2).ColumnWidth = 24  ' Set width to 25 'foto1
    otherSheet.Columns(3).ColumnWidth = 24
    otherSheet.Columns(4).ColumnWidth = 24
    otherSheet.Columns(5).ColumnWidth = 24
    otherSheet.Columns(6).ColumnWidth = 24 ' col width foto5
    
    otherSheet.Columns(10).ColumnWidth = 20
    otherSheet.Columns(11).ColumnWidth = 20
    otherSheet.Columns(12).ColumnWidth = 20
    otherSheet.Columns(13).ColumnWidth = 20
    otherSheet.Columns(14).ColumnWidth = 20
    otherSheet.Columns(15).ColumnWidth = 20
    
    

    ' Copy the content from cell (3,1) in hoja_main to cell (1,1) in otherSheet
    hoja_main.Cells(3, 1).Copy Destination:=otherSheet.Cells(1, 1) ' marca
    hoja_main.Cells(3, 3).Copy Destination:=otherSheet.Cells(1, 2) ' foto1
    hoja_main.Cells(3, 4).Copy Destination:=otherSheet.Cells(1, 3)
    hoja_main.Cells(3, 5).Copy Destination:=otherSheet.Cells(1, 4)
    hoja_main.Cells(3, 6).Copy Destination:=otherSheet.Cells(1, 5)
    hoja_main.Cells(3, 7).Copy Destination:=otherSheet.Cells(1, 6) ' foto5
    
    
    hoja_main.Cells(3, 21).Copy Destination:=otherSheet.Cells(1, 7) ' CTNS
    hoja_main.Cells(3, 22).Copy Destination:=otherSheet.Cells(1, 8) ' Cantidad
    otherSheet.Range("H1:I1").Merge
    otherSheet.Cells(1, 8).HorizontalAlignment = xlCenter
    
    hoja_main.Cells(3, 24).Copy Destination:=otherSheet.Cells(1, 10) ' Precio RMB
    
    hoja_main.Cells(3, 30).Copy Destination:=otherSheet.Cells(1, 11) ' Total
    hoja_main.Cells(3, 35).Copy Destination:=otherSheet.Cells(1, 12) ' CBM Caja
    hoja_main.Cells(3, 36).Copy Destination:=otherSheet.Cells(1, 13) ' CBM  total
    
    'copy only format
    hoja_main.Cells(3, 30).Copy
    otherSheet.Cells(1, 14).PasteSpecial Paste:=xlPasteFormats
    otherSheet.Cells(1, 15).PasteSpecial Paste:=xlPasteFormats
    otherSheet.Cells(1, 16).PasteSpecial Paste:=xlPasteFormats
    
    otherSheet.Cells(1, 14).Value = "COSTO"
    otherSheet.Cells(1, 15).Value = "SUGERIDO"
    
    ' Optionally, confirm the action is complete
    ' MsgBox "Column width adjusted and cell content copied successfully.", vbInformation
End Sub



Sub CopiarRecord(hoja_main As Worksheet, dst_sheet As Worksheet, row_orig As Integer, row_dst As Integer)
    
    dst_sheet.Rows(row_dst).rowHeight = 116
    
    'copy marca
    hoja_main.Cells(row_orig, 1).Copy Destination:=dst_sheet.Cells(row_dst, 1)
    
    
    'copy fotos
    hoja_main.Cells(row_orig, 3).Copy Destination:=dst_sheet.Cells(row_dst, 2) 'foto1
    hoja_main.Cells(row_orig, 4).Copy Destination:=dst_sheet.Cells(row_dst, 3)
    hoja_main.Cells(row_orig, 5).Copy Destination:=dst_sheet.Cells(row_dst, 4)
    hoja_main.Cells(row_orig, 6).Copy Destination:=dst_sheet.Cells(row_dst, 5)
    hoja_main.Cells(row_orig, 7).Copy Destination:=dst_sheet.Cells(row_dst, 6) 'foto5
    
    
    hoja_main.Cells(row_orig, 21).Copy Destination:=dst_sheet.Cells(row_dst, 7) 'CTNS
    
    hoja_main.Cells(row_orig, 22).Copy Destination:=dst_sheet.Cells(row_dst, 8) 'Cantidad 1 (formula)
    dst_sheet.Cells(row_dst, 8).Value = hoja_main.Cells(row_orig, 22).Value 'Cantidad 1
    
    hoja_main.Cells(row_orig, 23).Copy Destination:=dst_sheet.Cells(row_dst, 9) 'Cantidad 2
    hoja_main.Cells(row_orig, 24).Copy Destination:=dst_sheet.Cells(row_dst, 10) 'Precio RMB
    
    'Dim formulaTotalRMB As String
    'formulaTotalRMB = GenerateR1C1ProductFormula(row_dst, 8, row_dst, 10)
    hoja_main.Cells(row_orig, 30).Copy Destination:=dst_sheet.Cells(row_dst, 11) 'Total RMB (formula)
    dst_sheet.Cells(row_dst, 11).Value = hoja_main.Cells(row_orig, 30).Value 'Total RMB
    'dst_sheet.Cells(row_dst, 11).FormulaR1C1 = formulaTotalRMB 'Total RMB (formula)
    
    hoja_main.Cells(row_orig, 35).Copy Destination:=dst_sheet.Cells(row_dst, 12) 'CBM Caja
    
    'Dim formulaCBMTotal As String
    'formulaCBMTotal = GenerateR1C1ProductFormula(row_dst, 7, row_dst, 12)
    hoja_main.Cells(row_orig, 36).Copy Destination:=dst_sheet.Cells(row_dst, 13) 'CBM total (formula)
    dst_sheet.Cells(row_dst, 13).Value = hoja_main.Cells(row_orig, 36).Value 'CBM total
    'dst_sheet.Cells(row_dst, 13).FormulaR1C1 = "=R[row_dst]C[7] * R[row_dst]C[12]" 'CBM total
    
    
    
    
End Sub


' al final parece que es mejor no usar esta funcion
Function GenerateR1C1ProductFormula(startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer) As String
    ' This function generates an R1C1 notation string for the product of values in two specified cells.
    
    ' R1C1 notation for the starting cell
    Dim startCellR1C1 As String
    startCellR1C1 = "R" & startRow & "C" & startCol
    
    ' R1C1 notation for the ending cell
    Dim endCellR1C1 As String
    endCellR1C1 = "R" & endRow & "C" & endCol
    
    ' Construct the formula to return the product of the two cells
    Dim productFormula As String
    productFormula = "=" & startCellR1C1 & " * " & endCellR1C1
    
    ' Set the function name to the value you wish to return
    GenerateR1C1ProductFormula = productFormula
End Function



Sub MergeCellsInColumnA(startRow As Integer, endRow As Integer, ws As Worksheet)
    ' Check if the startRow is less than or equal to endRow
    If startRow <= endRow Then
        ' Specify the range to merge in column A (column 1) from startRow to endRow
        With ws
            .Range(.Cells(startRow, 1), .Cells(endRow, 1)).Merge
        End With
    Else
        ' Display a message box if startRow is greater than endRow
        MsgBox "The start row must be less than or equal to the end row.", vbExclamation
    End If
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Main()

    Dim hoja_main As Worksheet
    Set hoja_main = ThisWorkbook.Worksheets("Despacho")
    DeleteAllExceptDespacho
    
    Dim hoja_mdh As Worksheet
    Dim hoja_tin As Worksheet
    Dim hoja_like As Worksheet
    Dim hoja_alma As Worksheet
    
    Set hoja_mdh = CreateSheetIfNotExist("MDH")
    Set hoja_tin = CreateSheetIfNotExist("TIN")
    Set hoja_like = CreateSheetIfNotExist("LIKE")
    Set hoja_alma = CreateSheetIfNotExist("ALMA")
    
    
    'CleanWorkBook 'ensures that any worksheet not named as "Despacho" or "MDH" or "TIN" or "LIKE" or "ALMA" is going to be deleted
    
    Call SetupColTitles(hoja_main, hoja_mdh)
    Call SetupColTitles(hoja_main, hoja_tin)
    Call SetupColTitles(hoja_main, hoja_like)
    Call SetupColTitles(hoja_main, hoja_alma)
    
    
    Dim not_finished As Boolean
    not_finished = True
    
    
    ' indices para cada hoja
    Dim n As Integer ' indice hoja_main
    Dim m As Integer ' indice mdh
    Dim t As Integer
    Dim l As Integer
    Dim a As Integer ' indice alma sheet
    
    n = 4
    
    m = 2
    t = 2
    l = 2
    a = 2
    
    ' cellValue = ws.Cells(1, 1).Value  ' Cells(row number, column number)
    ' MsgBox "Cell (" & y & ", " & x & ") is completely empty."
    
    Dim nc As Integer
    
    
    While not_finished
        
        If IsEmpty(hoja_main.Cells(n, 1).Value) Then
            not_finished = False
            
        Else ' si la celda no esta vacia
            Dim bodega_name As String
            bodega_name = hoja_main.Cells(n, 1).Value
            'MsgBox bodega_name
            
            nc = CountMergedCells(hoja_main, n, 1)
        
            If nc > 1 Then ' si las celdas son merged
            
            Dim i As Integer
            ' Loop from n to n + nc
                For i = n To n + nc
                    Select Case bodega_name ' Copy Rows Case
                    Case "MDH"
                        Call CopiarRecord(hoja_main, hoja_mdh, i, m)
                        m = m + 1
                    Case "TIN"
                        Call CopiarRecord(hoja_main, hoja_tin, i, t)
                        t = t + 1
                    Case "LIKE"
                        Call CopiarRecord(hoja_main, hoja_like, i, l)
                        l = l + 1
                    Case "ALMA BEAUTY"
                        Call CopiarRecord(hoja_main, hoja_alma, i, a)
                        a = a + 1
                    End Select
                Next i
                
                Select Case bodega_name ' MergeCase Cells
                    Case "MDH"
                        Call MergeCellsInColumnA(m - nc - 1, m - 2, hoja_mdh)
                    Case "TIN"
                        Call MergeCellsInColumnA(t - nc - 1, t - 2, hoja_tin)
                    Case "LIKE"
                        Call MergeCellsInColumnA(l - nc - 1, l - 2, hoja_like)
                    Case "ALMA BEAUTY"
                        Call MergeCellsInColumnA(a - nc - 1, a - 2, hoja_alma)
                    End Select
                    
                n = n + nc
                
            Else ' si es una sola celda
                
                Select Case bodega_name ' LCase to make it case insensitive
                    Case "MDH"
                        Call CopiarRecord(hoja_main, hoja_mdh, n, m)
                        m = m + 1
                    Case "TIN"
                        Call CopiarRecord(hoja_main, hoja_tin, n, t)
                        t = t + 1
                    Case "LIKE"
                        Call CopiarRecord(hoja_main, hoja_like, n, l)
                        l = l + 1
                    Case "ALMA BEAUTY"
                        Call CopiarRecord(hoja_main, hoja_alma, n, a)
                        a = a + 1
                End Select
                
                n = n + 1
            End If
            
            
        End If
        
    Wend
    
    MsgBox "Count:  " & n
    
    'Call GetCellDimensions(hoja_main, 4, 1)
    'Call CopiarRecord(hoja_main, hoja_mdh, 4, 2)
    
End Sub


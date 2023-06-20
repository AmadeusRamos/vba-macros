Sub trim()

    
    For Each celda In Selection
    celda.Value = WorksheetFunction.trim(celda.Value)
    Next
    
        
End Sub

Sub comas()


    Set aa = Range("W2", Range("W2").End(xlDown))
    Set bb = Range("X2", Range("X2").End(xlDown))
    Set cc = Range("Y2", Range("Y2").End(xlDown))
    Set dd = Range("Z2", Range("Z2").End(xlDown))
    Set ee = Range("AA2", Range("AA2").End(xlDown))
        
    Union(aa, bb, cc, dd, ee).Select
    
    With Selection
    
    .Replace What:=",", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With
    
End Sub


Sub codcierr()


Dim cod_Cierr As String
Dim ultFila As Long
Dim cont As Long
Dim celda As Range

'COD_CIERR

ultFila = Range("A" & rows.Count).End(xlUp).Row

    For cont = 2 To ultFila
        cod_Cierr = Cells(cont, 13)
        
        If cod_Cierr = "-" Then
         Cells(cont, 13) = "SIN DATO"
         
        End If
    Next cont

End Sub

Sub homicidios()

Dim Rng As Range
Dim WorkRng As Range

On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    For Each Rng In WorkRng
        Rng.Value = VBA.UCase(Rng.Value)
    Next
    
    With Selection
    
    .Replace What:="á", Replacement:="a", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="é", Replacement:="e", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="í", Replacement:="i", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ó", Replacement:="o", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ú", Replacement:="u", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ü", Replacement:="u", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With
    
    
    


End Sub

Sub UCase()
'Upadateby20140701
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    Rng.Value = VBA.UCase(Rng.Value)
Next
End Sub


Sub accidentes()

    Dim celda As Range

'Transforma los datos almacenados en la columna ID de cadena de texto a numérico
    
    Range("A2", Range("A2").End(xlDown)).Select

    For Each celda In Selection
        celda.Value = CStr(celda)
    Next celda

'Ordena los datos de la hoja Coordenadas de menor a menor por el campo ID

    Range("A1").CurrentRegion.Sort Key1:=Range("A1"), Order1:=xlAscending, _
     Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'Se dividen los datos contenidos en el campo Coordenadas por el metodo tabulación y dividiendo por coma

    Columns("D:E").Insert Shift:=xlToRight
    Range("C2", Range("C1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)

'Se copian los datos del rango de columnas "D:I" y se pegan en la columnas "BB:BG"
    
    ActiveSheet.Range("D2", Range("D2").End(xlDown)).Select
    Selection.resize(, 6).Copy
    Sheets("ACCIDENTES").Select
    Range("BB2").Select
    Selection.PasteSpecial Paste:=xlPasteValues

'Se ordenan las columnas de acuerdo a la información por campo

    Range("BE2", Range("BE2").End(xlDown)).Cut Range("BA2")
    Range("BD2", Range("BD1048576").End(xlUp)).Cut Range("AZ2")
    Range("BF2", Range("BF1048576").End(xlUp)).Cut Range("AU2")
    Range("BG2", Range("BG1048576").End(xlUp)).Cut Range("AV2")
    Range("BC2", Range("BC1048576").End(xlUp)).Cut Range("BD2")
    Range("BB2", Range("BB1048576").End(xlUp)).Cut Range("BC2")
    Range("BC2").Select

End Sub

Sub textoANumero()

Dim celda As Range

Range("A2", Range("A2").End(xlDown)).Select

For Each celda In Selection
    celda.Value = CStr(celda)
Next celda


End Sub


Sub resize()

Dim LR As Long
Dim LC As Long
    
    Worksheets("COORDENADAS").Select
    
    LR = Cells(rows.Count, 1).End(xlUp).Row
    LC = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Cells(1, 1).resize(LR, LC).Select
    
    
'Dim rows As Long, numcolumns As Integer

'numrows = Selection.rows.Count
'numcolumns = Selection.Columns.Count
'Selection.resize(numrows + 10, numcolumns + 5).Select



  End Sub
  
  Option Explicit
Sub RecorrerCeldas()

Dim celda As Range

For Each celda In Selection

    If Not VBA.IsNumeric(celda) Then
        celda.Font.Bold = True
    End If

Next celda

End Sub


Sub ValidarCelda()

Dim celda As Range

For Each celda In Range("A2:BK14")

    If VBA.IsNumeric(celda) = True Then
        celda.Interior.Color = VBA.vbGreen
                  
    End If

Next celda

For Each celda In Range("A2:BK14")

    If Not VBA.IsNumeric(celda) Then
        celda.Font.Bold = True
    
    End If

Next celda

End Sub

Sub eliminarFilas()
         
'Esta macro elimina toda la fila
'De las celdas seleccionadas
         
Application.ScreenUpdating = False
         
   Selection.EntireRow.Delete
   
Application.ScreenUpdating = True
    
End Sub

Sub seleccionarColumna()

    ' Esta Macro solamente acepta datos contiguos, se detiene en celdas en blanco
    'Range("AX2", Range("AX2").End(xlDown)).Select
    ' Esta Macro puede elegir los datos de una columna aun no siendo contiguos
    'Range("AX2", Range("AX1048576").End(xlUp)).Select
    ' Esta Macro resalta toda el área de trabajo en uso
    Selection.CurrentRegion.Select
    

End Sub

Sub experimentos()

Dim celda As Range

Selection.CurrentRegion.Select

With Selection

.Replace What:="" & Chr(10) & "", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns
        
End With

For Each celda In Selection
celda.Value = LTrim(celda.Value)
celda.Value = RTrim(celda.Value)
celda.Value = WorksheetFunction.trim(celda.Value)
Next

End Sub
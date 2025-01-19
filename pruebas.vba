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

Sub cambiarSeleccion()


ActiveSheet.Range("A2", Range("A2").End(xlDown)).Select
Selection.resize(, 46).Select


End Sub

Sub Renumbering()
'https://es.extendoffice.com/documents/excel/2209-excel-auto-number-after-filter.html#:~:text=1%201.%20Mantenga%20pulsado%20el%20ALT%20%2B%20F11,o%20renumerar%20despu%C3%A9s%20del%20filtro%20...%20M%C3%A1s%20elementos
    'Updateby Extendoffice
    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    Set WorkRng = WorkRng.Columns(1).SpecialCells(xlCellTypeVisible)
    xIndex = 1
    For Each rng In WorkRng
        rng.Value = xIndex
        xIndex = xIndex + 1
    Next
End Sub

Sub AñadirSeries()
'https://www.desafiandoexcel.com/macros-de-excel/numeros-seriales-o-secuenciales-en-excel-macro/#:~:text=La%20siguiente%20macro%20nos%20permite%20generar%20r%C3%A1pidamente%20una,debemos%20configurar%20el%20%C3%BAltimo%20valor%20para%20dicha%20serie.
Dim i As Integer
On Error GoTo Last
i = InputBox("Ingresar último valor", "Números seriales")
For i = 1 To i
ActiveCell.Value = i
ActiveCell.Offset(1, 0).Activate
Next i
Last: Exit Sub
End Sub

Sub NumeracionAutomatica()

Dim i As Integer

i = 1

Do While Not IsEmpty(Cells(i, 1))
Cells(i, 1).Value = i
i = i + 1

Loop

End Sub

Sub AutoGeneratedSerialNumber()
'https://www.exceldemy.com/auto-generate-serial-number-in-excel-vba/
    For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
            Cells(m, "A").Value = m - 1
        End If
    Next m
End Sub

Sub cambiarSeleccion()


ActiveSheet.Range("A2", Range("A2").End(xlDown)).Select
Selection.resize(, 46).Select


End Sub

Sub Renumbering()
'https://es.extendoffice.com/documents/excel/2209-excel-auto-number-after-filter.html#:~:text=1%201.%20Mantenga%20pulsado%20el%20ALT%20%2B%20F11,o%20renumerar%20despu%C3%A9s%20del%20filtro%20...%20M%C3%A1s%20elementos
    'Updateby Extendoffice
    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    Set WorkRng = WorkRng.Columns(1).SpecialCells(xlCellTypeVisible)
    xIndex = 1
    For Each rng In WorkRng
        rng.Value = xIndex
        xIndex = xIndex + 1
    Next
End Sub

Sub AñadirSeries()
'https://www.desafiandoexcel.com/macros-de-excel/numeros-seriales-o-secuenciales-en-excel-macro/#:~:text=La%20siguiente%20macro%20nos%20permite%20generar%20r%C3%A1pidamente%20una,debemos%20configurar%20el%20%C3%BAltimo%20valor%20para%20dicha%20serie.
Dim i As Integer
On Error GoTo Last
i = InputBox("Ingresar último valor", "Números seriales")
For i = 1 To i
ActiveCell.Value = i
ActiveCell.Offset(1, 0).Activate
Next i
Last: Exit Sub
End Sub

Sub NumeracionAutomatica()

Dim i As Integer

i = 1

Do While Not IsEmpty(Cells(i, 1))
Cells(i, 1).Value = i
i = i + 1

Loop

End Sub

Sub AutoGeneratedSerialNumber()
'https://www.exceldemy.com/auto-generate-serial-number-in-excel-vba/
    For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
            Cells(m, "A").Value = m - 1
        End If
    Next m
End Sub

Sub recorte()

Range("S:S").Value = trim(Range("S:S"))

End Sub


'https://exceltotal.com/como-quitar-acentos-en-excel/

Function QUITARACENTOS(cadena As String) As String
Dim posicion As Long
Const conAcento As String = "áéíóúÁÉÍÓÚ"
Const sinAcento As String = "aeiouAEIOU"

For i = 1 To Len(conAcento)
    cadena = Replace(cadena, Mid(conAcento, i, 1), Mid(sinAcento, i, 1))
Next i

QUITARACENTOS = cadena

End Function

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
    
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    LC = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Cells(1, 1).resize(LR, LC).Select
    
    
'Dim rows As Long, numcolumns As Integer

'numrows = Selection.rows.Count
'numcolumns = Selection.Columns.Count
'Selection.resize(numrows + 10, numcolumns + 5).Select



End Sub

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

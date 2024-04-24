Sub eliminarLibroStrack()

Dim Ruta As String, wb As Workbook

Ruta = "C:\Test\accidentes\STRACK.xlsx"

If Dir(Ruta) = "" Then
MsgBox "Archivo no encontrado.:(": Exit Sub
End If

On Error Resume Next
Set wb = Workbooks(Dir(Ruta))

If Not wb Is Nothing Then
MsgBox "No se puede borrar archivo abierto.:("
    Else
        Kill Ruta
End If

Set wb = Nothing

End Sub

Sub moverLibroStrack()
' https://exceloffthegrid.com/vba-code-to-copy-move-delete-and-manage-files/

    Name "C:\Users\emmanuel.ramos\Desktop\STRACK.xlsx" As "C:\Test\accidentes\STRACK.xlsx"

End Sub

Sub abrirLibroStrack()
'https://www.gerencie.com/macro-para-abrir-un-libro-de-excel-desde-otro-libro.html

    Workbooks.Open "C:\Test\accidentes\STRACK.xlsx"

End Sub

Sub cambiarNombreHoja()

Workbooks("STRACK.xlsx").Activate
ActiveSheet.Name = "reporte"

End Sub

Sub eliminarND()

Application.ScreenUpdating = False

'Esta macro elimina los N/D y No Calificados
    Range("A1:AQ1").AutoFilter field:=6, _
        Criteria1:=Array("_NO CALIFICADO_", "#N/D"), Operator:=xlFilterValues
        
    With ActiveSheet.AutoFilter.Range
    .Offset(1, 0).resize(.Rows.Count - 1).Select
    End With

    Selection.EntireRow.Delete
    
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
Application.ScreenUpdating = True
    
End Sub

Sub abrirLibroLimpieza()
'https://foro.todoexcel.com/threads/abrir-con-vba-libro-protegido.7777/

    Workbooks.Open "C:\Test\accidentes\04 limpiezaGral_4.xlsm", Password:="1357"

End Sub

Sub limpiarRango()
'https://www.vbatotal.com/leccion-21-seleccionar-una-hoja-o-un-libro-automaticamente-con-vba/

    Workbooks("04 limpiezaGral_4.xlsm").Activate

    Range("A2", Range("A1048576").End(xlUp)).Select
        Selection.resize(, 43).ClearContents

End Sub

Sub copiarStrackLimpieza()
'Esta macro va a copiar los datos desde el libro STRACK al libro 04 limpiezaGral_4

'Application.ScreenUpdating = False
    
    Workbooks("STRACK.xlsx").Activate
    
    'Formato del espacio de trabajo
    Range("A2").CurrentRegion.Select
    With Selection

    .Font.Name = "Arial"
    .Font.Size = 9
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Orientation = 0
    .IndentLevel = 0
    .ReadingOrder = xlContext

    End With
    
    Selection.Range("A2", Range("A1048576").End(xlUp)).Select
    Selection.resize(, 43).Copy
    Workbooks("04 limpiezaGral_4.xlsm").Sheets("geo").Activate
    Range("A2").Select
    ActiveSheet.Paste
    
'Application.ScreenUpdating = True
Application.CutCopyMode = False

End Sub

Sub copiarFormulas()

ActiveWorkbook.Sheets("geo").Activate

Range("AR2").AutoFill Range("AR2:AR" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AS2").AutoFill Range("AS2:AS" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AT2").AutoFill Range("AT2:AT" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AU2").AutoFill Range("AU2:AU" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AV2").AutoFill Range("AV2:AV" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AW2").AutoFill Range("AW2:AW" & Range("B" & Rows.Count).End(xlUp).Row)
Range("AR2").Select

End Sub

Sub cerrarLibroStrack()

Application.ScreenUpdating = False

    Workbooks("STRACK.xlsx").Activate
    Range("A2").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close

Application.ScreenUpdating = True

End Sub

Sub copiarDai2()
'Esta macro copia los datos del libro 04 limpiezaGral_4 de la columna AR a la columna E
'Esto es de la columna DAI2 a la columna DAI

Application.ScreenUpdating = False

    Workbooks("04 limpiezaGral_4.xlsm").Sheets("geo").Activate
    Range("AR2", Range("AR1048576").End(xlUp)).Copy
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
Application.ScreenUpdating = True
Application.CutCopyMode = False
    
End Sub

Sub suprimirDuplicadosND()

    Dim lastRow As Long, FirstRow As Long
    Dim Row As Long
    
Application.ScreenUpdating = False
    
'Esta macro elimina los errores #N/D y duplicados de las columnas DAI2, municipio y
'Folios repetidos del viernes
    
    With ActiveSheet
        'Definir la primera y la última fila
        FirstRow = 2
        lastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
    
        'Bucle a través de filas (de abajo a arriba) en las columnas AR AU y AV
        'Para detectar los errores #N/D en las hojas, el tipo a usar es Text no Value
        'Explicación en
        'https://es.stackoverflow.com/questions/538707/eliminar-filas-n-d-con-una-macro
        For Row = lastRow To FirstRow Step -1
            If .Range("AR" & Row).Text = "#N/D" Then
                .Range("AR" & Row).EntireRow.Delete
            ElseIf .Range("AU" & Row).Text = "#N/D" Then
                    .Range("AU" & Row).EntireRow.Delete
            ElseIf .Range("AV" & Row).Text <> "#N/D" Then
                    .Range("AV" & Row).EntireRow.Delete
            End If
        Next Row
    End With
    
Application.ScreenUpdating = True
        
        Range("AR2").Select
        
End Sub

Sub limpiarHojas()

    Workbooks("04 limpiezaGral_4.xlsm").Activate
    Application.ScreenUpdating = False
    Sheets("DAI").Select
    Range("A2", Range("A1048576").End(xlUp)).Select
        Selection.EntireRow.Delete
            Range("A2").Select
                Sheets("seguridad").Select
                    Range("A2", Range("A1048576").End(xlUp)).Select
                        Selection.EntireRow.Delete
                            Range("A2").Select
                                Sheets("accidentes").Select
                                    Range("A2", Range("A1048576").End(xlUp)).Select
                                        Selection.EntireRow.Delete
                                            Range("A2").Select
    Application.ScreenUpdating = True

End Sub

Sub ordenaPorMunicipio()

Workbooks("04 limpiezaGral_4.xlsm").Activate
Sheets("geo").Select

Application.ScreenUpdating = False

'Ordena la columna W
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "W2", Range("W2").End(xlDown)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
'Ordena la columna Z
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "Z2", Range("Z2").End(xlDown)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Ordena la columna V
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "V2", Range("V2").End(xlDown)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
    
    Range("AY1").Select
End Sub

Sub filasDAI()

Workbooks("04 limpiezaGral_4.xlsm").Activate
Sheets("geo").Select

Application.ScreenUpdating = False

'Copia las celdas con criterio DAI en la columna AW

'inicializo la variable j
j = 2

    'comienzo el bucle
    ultFila = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultFila
        
        'activo la hoja donde están mis datos
        Sheets("geo").Activate
            
            'compruebo que el valor de la fecha es mayor que 30
            If Cells(i, "AW").Text = "DAI" Then
                'copio la fila entera
                Range(Cells(i, "A"), Cells(i, "AT")).Copy
                'selecciono la hoja donde quiero pegar y después la celda
                Sheets("DAI").Activate
                    Cells(j, "A").Select
                'pego la fila que hemos copiado
                ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats
                'aumento la variable j para que vaya a la siguiente fila de la hoja filtros
                'cuando encuentre una nueva fila que cumple con la condición de edad
                j = j + 1
            End If
    Next
    
    'Elegir la hoja de trabajo y ordenar las columnas
    Sheets("DAI").Select
    Range("AP2", Range("AP1048576").End(xlUp)).Clear
    Range("AQ2", Range("AQ1048576").End(xlUp)).Clear
    Range("AR2", Range("AR2").End(xlDown)).Clear
    Range("AS2", Range("AS2").End(xlDown)).Cut Range("BH2")
    Range("AT2", Range("AT2").End(xlDown)).Cut Range("BK2")
    
    'Numeración consecutiva de la columna ID
    For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
         Cells(m, "A").Value = m - 1
        End If
    Next m
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    
    Range("A2").Select

End Sub

Sub filasSeguridad()

Workbooks("04 limpiezaGral_4.xlsm").Activate
Sheets("geo").Select

Application.ScreenUpdating = False

'Copia las celdas con criterio RESTO en la columna AW

'inicializo la variable j
j = 2

    'comienzo el bucle
    ultFila = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultFila
        
        'activo la hoja donde están mis datos
        Sheets("geo").Activate
            
            'compruebo que el valor de la fecha es mayor que 30
            If Cells(i, "AW").Text = "RESTO" Then
                'copio la fila entera
                Range(Cells(i, "A"), Cells(i, "AT")).Copy
                'selecciono la hoja donde quiero pegar y después la celda
                Sheets("seguridad").Activate
                    Cells(j, "A").Select
                'pego la fila que hemos copiado
                ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats
                'aumento la variable j para que vaya a la siguiente fila de la hoja filtros
                'cuando encuentre una nueva fila que cumple con la condición de edad
                j = j + 1
            End If
    Next
    
    'Elegir la hoja de trabajo y ordenar las columnas
    Sheets("seguridad").Select
    Range("AP2", Range("AP1048576").End(xlUp)).Clear
    Range("AQ2", Range("AQ1048576").End(xlUp)).Clear
    Range("AR2", Range("AR2").End(xlDown)).Clear
    Range("AS2", Range("AS2").End(xlDown)).Cut Range("BH2")
    Range("AT2", Range("AT2").End(xlDown)).Cut Range("BK2")
    
    'Numeración consecutiva de la columna ID
    For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
         Cells(m, "A").Value = m - 1
        End If
    Next m
  
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    
    Range("A2").Select

End Sub

Sub filasAccidentes()

Workbooks("04 limpiezaGral_4.xlsm").Activate
Sheets("geo").Select

Application.ScreenUpdating = False

'Copia las celdas con criterio ACCIDENTES en la columna AW

'inicializo la variable j
j = 2

    'comienzo el bucle
    ultFila = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultFila
        
        'activo la hoja donde están mis datos
        Sheets("geo").Activate
            
            'compruebo que el valor de la fecha es mayor que 30
            If Cells(i, "AW").Text = "ACCIDENTES" Then
                'copio la fila entera
                Range(Cells(i, "A"), Cells(i, "AT")).Copy
                'selecciono la hoja donde quiero pegar y después la celda
                Sheets("accidentes").Activate
                    Cells(j, "A").Select
                'pego la fila que hemos copiado
                ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats
                'aumento la variable j para que vaya a la siguiente fila de la hoja filtros
                'cuando encuentre una nueva fila que cumple con la condición de edad
                j = j + 1
            End If
    Next
    
    'Elegir la hoja de trabajo y ordenar las columnas
    Sheets("accidentes").Select
    Range("AP2", Range("AP1048576").End(xlUp)).Clear
    Range("AQ2", Range("AQ1048576").End(xlUp)).Clear
    Range("AR2", Range("AR2").End(xlDown)).Clear
    Range("AS2", Range("AS2").End(xlDown)).Cut Range("BH2")
    Range("AT2", Range("AT2").End(xlDown)).Cut Range("BK2")
    
    'Numeración consecutiva de la columna ID
    For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
         Cells(m, "A").Value = m - 1
        End If
    Next m
 
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    
    Range("A2").Select

End Sub

Sub abrirLibroGeneral()

Workbooks.Open "C:\Test\accidentes\general.xlsx"
'Workbooks.Open "C:\Users\Naty\Desktop\Test\general.xlsx"

End Sub

Sub limpiarLibroGeneral()

    Workbooks("general.xlsx").Activate
    Application.ScreenUpdating = False
    Sheets("DAI").Select
        Range("A2", Range("A1048576").End(xlUp)).Select
            Selection.EntireRow.Delete
                Sheets("seguridad").Select
                    Range("A2", Range("A1048576").End(xlUp)).Select
                        Selection.EntireRow.Delete
                            Sheets("accidentes").Select
                                Range("A2", Range("A1048576").End(xlUp)).Select
                                    Selection.EntireRow.Delete
    Application.ScreenUpdating = True

End Sub

Sub deLimpieza_a_General()
'Esta macro copia los datos contenidos en el libro de limpieza y los copia en el libro general
'El libro general es el que se comparte para iniciar la georreferencia
    
    Workbooks("04 limpiezaGral_4.xlsm").Activate
    Application.ScreenUpdating = False
    Sheets("DAI").Select
    Selection.Range("A1", Range("A1048576").End(xlUp)).Select
    Selection.resize(, 63).Copy
    Workbooks("general.xlsx").Activate
    Sheets("DAI").Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    
    Workbooks("04 limpiezaGral_4.xlsm").Activate
    Sheets("seguridad").Select
    Selection.Range("A1", Range("A1048576").End(xlUp)).Select
    Selection.resize(, 63).Copy
    Workbooks("general.xlsx").Activate
    Sheets("seguridad").Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    
    Workbooks("04 limpiezaGral_4.xlsm").Activate
    Sheets("accidentes").Select
    Selection.Range("A1", Range("A1048576").End(xlUp)).Select
    Selection.resize(, 63).Copy
    Workbooks("general.xlsx").Activate
    Sheets("accidentes").Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    
End Sub

Sub cerrarLibroLimpieza()

Application.ScreenUpdating = False

    Workbooks("04 limpiezaGral_4.xlsm").Sheets("geo").Activate
    Range("A2").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close

Application.ScreenUpdating = True

End Sub

Sub cerrarLibroGeneral()

Application.ScreenUpdating = False

    Workbooks("general.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close

Application.ScreenUpdating = True

End Sub
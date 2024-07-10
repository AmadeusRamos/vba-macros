Sub celdasDAI()
'Esta macro se ejecuta por medio de un botón creado exclusivamente para los datos de DAI
'Procedimiento
'Copia todos los datos de la hoja "geo" y los pega en la hoja "DAI"
'En la hoja destino, ordena la numeración de 1 a n para cada folio por trabajar
'Limpia los datos de las columnas hom_tot, hom_hombr y hom_muj
'Actualiza los datos de las columnas aux y carto

'Definir objetos a utilizar

    Dim wsOrigen As Excel.Worksheet, _
        wsDestino As Excel.Worksheet, _
        rngOrigen As Excel.Range, _
        rngDestino As Excel.Range

    Application.ScreenUpdating = False

'Indicar las hojas de origen y destino
    Set wsOrigen = Worksheets("geo")
    Set wsDestino = Worksheets("DAI")

'Indicar la celda de origen y destino
    Const celdaOrigen = "A2"
    Const celdaDestino = "A2"
    
'Inicializar los rangos de origen y destino
    Set rngOrigen = wsOrigen.Range(celdaOrigen)
    Set rngDestino = wsDestino.Range(celdaDestino)
    
'Seleccionar rango de celdas origen
    rngOrigen.Select
    ActiveSheet.Range("A2", Range("A2").End(xlDown)).Select
    Selection.resize(, 46).Copy
    'Selection.Copy
    
'Pegar datos en celda destino
    rngDestino.PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    
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
    
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    Application.ScreenUpdating = True
    
    Range("A2").Select
    Sheets("geo").Select
    Range("AW1").Select
    
End Sub

Sub celdasSeguridad()
'Esta macro se ejecuta por medio de un botón creado exclusivamente para los datos de Seguridad
'Procedimiento
'Copia todos los datos de la hoja "geo" y los pega en la hoja "seguridad"
'En la hoja destino, ordena la numeración de 1 a n para cada folio por trabajar
'Limpia los datos de las columnas hom_tot, hom_hombr y hom_muj
'Actualiza los datos de las columnas aux y carto

'Definir objetos a utilizar

    Dim wsOrigen As Excel.Worksheet, _
        wsDestino As Excel.Worksheet, _
        rngOrigen As Excel.Range, _
        rngDestino As Excel.Range
    
    Application.ScreenUpdating = False
        
'Indicar las hojas de origen y destino
    Set wsOrigen = Worksheets("geo")
    Set wsDestino = Worksheets("seguridad")

'Indicar la celda de origen y destino
    Const celdaOrigen = "A2"
    Const celdaDestino = "A2"
    
'Inicializar los rangos de origen y destino
    Set rngOrigen = wsOrigen.Range(celdaOrigen)
    Set rngDestino = wsDestino.Range(celdaDestino)
    
'Seleccionar rango de celdas origen
    rngOrigen.Select
    ActiveSheet.Range("A2", Range("A2").End(xlDown)).Select
    Selection.resize(, 46).Copy
    'Selection.Copy
    
'Pegar datos en celda destino
    rngDestino.PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    
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
    
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    Application.ScreenUpdating = True
    
    Range("A2").Select
    Sheets("geo").Select
    Range("AW1").Select
    
End Sub

Sub celdasAccidentes()
'Esta macro se ejecuta por medio de un botón creado exclusivamente para los datos de Accidentes
'Procedimiento
'Copia todos los datos de la hoja "geo" y los pega en la hoja "accidentes"
'En la hoja destino, ordena la numeración de 1 a n para cada folio por trabajar
'Limpia los datos de las columnas hom_tot, hom_hombr y hom_muj
'Actualiza los datos de las columnas aux y carto

'Definir objetos a utilizar

    Dim wsOrigen As Excel.Worksheet, _
        wsDestino As Excel.Worksheet, _
        rngOrigen As Excel.Range, _
        rngDestino As Excel.Range
        
        
    Application.ScreenUpdating = False
        
'Indicar las hojas de origen y destino
    Set wsOrigen = Worksheets("geo")
    Set wsDestino = Worksheets("accidentes")

'Indicar la celda de origen y destino
    Const celdaOrigen = "A2"
    Const celdaDestino = "A2"
    
'Inicializar los rangos de origen y destino
    Set rngOrigen = wsOrigen.Range(celdaOrigen)
    Set rngDestino = wsDestino.Range(celdaDestino)
    
'Seleccionar rango de celdas origen
    rngOrigen.Select
    ActiveSheet.Range("A2", Range("A2").End(xlDown)).Select
    Selection.resize(, 46).Copy
    'Selection.Copy
    
'Pegar datos en celda destino
    rngDestino.PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    
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
    
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    Application.ScreenUpdating = True
    
    Range("A2").Select
    Sheets("geo").Select
    Range("AW1").Select
    
End Sub

Sub ordenarColumnas()
'Macro para ordenar los datos de la hoja geo del Libro 03 limpiezaGral_3.22
'Procedimiento
'Ordena de la A a la Z cada columna, iniciando con direccion, colonia y termina con municipio
'De esta forma los datos están en el orden normalizado para el flujo de trabajo

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

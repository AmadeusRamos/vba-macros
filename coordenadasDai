Sub coordenadasDAI()
'Esta macro realiza la division de las coordenadas, ordena de menor a mayor los folios georreferenciados, copia los datos de municipio, colonia, clave punteo, analista, latitud y longitud, pegando dichos datos en la hoja llamada coordenadas
'Esta macro solamente sirve para los archivos de DAI

    Dim celda As Range
    
    Application.ScreenUpdating = False
    
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
    Sheets("DAI").Select
    Range("BB2").Select
    Selection.PasteSpecial Paste:=xlPasteValues

'Se ordenan las columnas de acuerdo a la información por campo

    Range("BE2", Range("BE2").End(xlDown)).Cut Range("BA2")
    Range("BD2", Range("BD1048576").End(xlUp)).Cut Range("AZ2")
    Range("BF2", Range("BF1048576").End(xlUp)).Cut Range("AU2")
    Range("BG2", Range("BG1048576").End(xlUp)).Cut Range("AV2")
    Range("BC2", Range("BC1048576").End(xlUp)).Cut Range("BD2")
    Range("BB2", Range("BB1048576").End(xlUp)).Cut Range("BC2")
    
    Application.ScreenUpdating = True
    
    Range("BC2").Select

End Sub

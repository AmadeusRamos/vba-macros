Sub limpiarDatos089()

Dim wsCopia As Worksheet
Dim wsDestino As Worksheet
Dim copyLastRow As Long
Dim destLastRow As Long

'Esta macro ayuda a minimizar los tiempos de copiar y pegar datos al libro de comprobación
'Después de usar esta macro se procede a realizar la unión espacial en QGIS utlizando el modelo creado para tal fin
'Usar el archivo llamado 089_5.0.xlsx

    Application.DisplayAlerts = False

    'Establecer los libros a trabajar
    Set wsCopia = Workbooks("089_5.0.xlsx").Worksheets("Segmento")
    Set wsDestino = Workbooks.Open(ActiveWorkbook.Path & "\comprobacion.csv").Worksheets("comprobacion")
    
        '1. Encontrar la última fila con datos en el rango de origen en la columna A
        copyLastRow = wsCopia.Cells(wsCopia.Rows.Count, 1).End(xlUp).Row
        
        '2. Encontrar la última fila en blanco en el rango de destino sobre la columna A
        'La propiedad Offset se moverá 1 celda abajo
        destLastRow = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Offset(1).Row
        
        '3. limpiar el contenido de las celdas del libro destino
        wsDestino.Range("A2", "D" & destLastRow).ClearContents
        
        '4. Copiar y pegar los datos
        wsCopia.Range("A2", "C" & copyLastRow).Copy wsDestino.Range("A2")
        
        '5. Separar las coordenadas x e y
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
        Range("C2", Range("C1048576").End(xlUp)).Select
        With Selection
            .TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
            textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
        End With
        Selection.Delete Shift:=xlToLeft
        Application.DisplayAlerts = True
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close

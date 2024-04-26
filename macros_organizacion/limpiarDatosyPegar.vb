Sub limpiarDatosyPegar()

'https://www.excelcampus.com/vba/copy-paste-another-workbook/
'https://ayudaexcel.com/foro/topic/44179-rango-celdas-sin-conocer-el-n%C3%BAmero-exacto/
'https://www.excel-avanzado.com/2791/identificar-la-ultima-fila-en-uso-con-vba.html/comment-page-2
'https://es.wikibooks.org/wiki/Seleccionar_o_referenciar_celdas_de_Excel_mediante_VBA
'https://www.automateexcel.com/es/vba/variables-de-objetos-de-rango/
'https://www.google.com/search?client=firefox-b-d&q=metodo+rows.count+vba
'https://www.automateexcel.com/es/vba/abrir-cerrar-libro-de-trabajo/

Dim wsCopia As Worksheet
Dim wsDestino As Worksheet
Dim copyLastRow As Long
Dim destLastRow As Long

'Esta macro ayuda a pasar los datos finalizados de la geo al libro de comprobación
'Minimiza el tiempo de copiar y pegar.
'Después de usar esta macro se procede a realizar la unión espacial en QGIS utlizando el modelo creado para tal fin
'Usar el archivo llamado GEO.xlsx

    Application.DisplayAlerts = False

    'Establecer los libros a trabajar
    Set wsCopia = Workbooks("GEO.xlsx").Worksheets("Segmento")
    Set wsDestino = Workbooks.Open(ActiveWorkbook.Path & "\comprobacion.csv").Worksheets("comprobacion")
    
        '1. Encontrar la última fila con datos en el rango de origen en la columna A
        copyLastRow = wsCopia.Cells(wsCopia.Rows.Count, 1).End(xlUp).Row
        
        '2. Encontrar la última fila en blanco en el rango de detino sobre la columna A
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

End Sub
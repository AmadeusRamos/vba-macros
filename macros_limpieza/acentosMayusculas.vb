Sub acentosMayus()

'Ayuda a transformar cualquier rango a trabajar en acentos y sin mayúsculas

Application.ScreenUpdating = False

Dim rng As Range
Dim WorkRng As Range
Dim celda As Range

'La macro sin la siguiente línea transforma el formato de fecha cambiando el mes al día
'Si tenemos la fecha 05/06/2024 la pasa a 06/05/2024
'Pero con la siguiente línea forzamos el formato al modo inglés y lo cambia al modo que usamos
'Se debe elegir la columna objetivo, este ejemplo sirve para el archivo de infracciones
'https://es.stackoverflow.com/questions/213726/problema-de-conversion-de-fecha-utilizando-vba-de-excel
Range("D:D").NumberFormat = "MM/DD/YYYY"

Selection.CurrentRegion.Select

    With Selection

        .Replace What:="á", Replacement:="a", LookAt:=xlPart, SearchOrder:=xlByColumns
        .Replace What:="é", Replacement:="e", LookAt:=xlPart, SearchOrder:=xlByColumns
        .Replace What:="í", Replacement:="i", LookAt:=xlPart, SearchOrder:=xlByColumns
        .Replace What:="ó", Replacement:="o", LookAt:=xlPart, SearchOrder:=xlByColumns
        .Replace What:="ú", Replacement:="u", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With
    
    
    'Convierte en mayúsculas todo el rango en uso
    
    On Error Resume Next
    Set WorkRng = Application.Selection
    For Each rng In WorkRng
        rng.Value = VBA.UCase(rng.Value)
    Next
    
    With Selection

    .Font.Name = "Arial"
    .Font.Size = 10
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Orientation = 0
    .IndentLevel = 0
    .ReadingOrder = xlContext
    .RowHeight = 15

    End With
    
    Application.ScreenUpdating = True

    Range("A2").Select
    
End Sub
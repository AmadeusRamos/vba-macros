Option Explicit

Sub cartoStrack()

'Limpieza de Strack
'Esta acción elimina los bordes de las celdas
'Tambien elimina saltos de página y acentos
'Es importante contar 43 columnas, no más no menos

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim rng As Range
Dim WorkRng As Range
Dim celda As Range
Dim a As Range, b As Range, c As Range, d As Range, e As Range, f As Range
Dim g As Range, h As Range, i As Range, j As Range, k As Range, l As Range
Dim m As Range, n As Range, o As Range, p As Range, q As Range, r As Range
Dim s As Range, t As Range, u As Range, v As Range, w As Range, x As Range
Dim aa As Range, bb As Range, cc As Range, dd As Range, ee As Range

Selection.CurrentRegion.Select

    With Selection
    
    .Font.Bold = False
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
    .Replace What:="" & Chr(10) & "", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="á", Replacement:="a", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="é", Replacement:="e", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="í", Replacement:="i", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ó", Replacement:="o", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ú", Replacement:="u", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="à", Replacement:="a", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="è", Replacement:="e", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ì", Replacement:="i", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ò", Replacement:="o", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="ù", Replacement:="u", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="  ,  , ", Replacement:=", ", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:=" ,  , ", Replacement:=", ", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:=" ,, ", Replacement:=", ", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="/ / / / /", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="/ / / /", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="/ / /", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="/ /", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:=" / ", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="/////", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="////", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="///", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns
    .Replace What:="//", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByColumns

    End With
  
'Únicamente aplicar a las columnas "V:AO"
    Range("V2", Range("V2").End(xlDown).End(xlToRight)).Select

    With Selection
    
    .Replace What:="•", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="$", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="|", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="°", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="!", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="¡", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="""", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="#", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="%", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="&", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="(", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:=")", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="=", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="~*", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="~?", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="~¿", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="'", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="+", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="{", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="}", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="[", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="]", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="<", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:=">", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    .Replace What:="`", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
        
    End With
    
'Elimina las comas del rango de columnas W:AA

    Set aa = Range("W2", Range("W2").End(xlDown))
    Set bb = Range("X2", Range("X2").End(xlDown))
    Set cc = Range("Y2", Range("Y2").End(xlDown))
    Set dd = Range("Z2", Range("Z2").End(xlDown))
    Set ee = Range("AA2", Range("AA2").End(xlDown))
        
    Union(aa, bb, cc, dd, ee).Select
    
    With Selection
    
    .Replace What:=",", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With

'Convierte en mayúsculas todo el rango en uso
    
    On Error Resume Next
    Selection.CurrentRegion.Select
    Set WorkRng = Application.Selection
    For Each rng In WorkRng
        rng.Value = VBA.UCase(rng.Value)
    Next

'Se eliminan los guiones medios del rango AD:AF y AH:AO

Range("AD2", Range("AD2").End(xlDown).End(xlToRight)).Select
  
    With Selection
    
    .Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With

'Quita los espacios de los rangos D:N, V:AB Y AJ:AO
'Rango D:N
    Set a = Range("D2", Range("D2").End(xlDown))
    Set b = Range("E2", Range("E2").End(xlDown))
    Set c = Range("F2", Range("F2").End(xlDown))
    Set d = Range("G2", Range("G2").End(xlDown))
    Set e = Range("H2", Range("H2").End(xlDown))
    Set f = Range("I2", Range("I2").End(xlDown))
    Set g = Range("J2", Range("J2").End(xlDown))
    Set h = Range("K2", Range("K2").End(xlDown))
    Set i = Range("L2", Range("L2").End(xlDown))
    Set j = Range("M2", Range("M2").End(xlDown))
    Set k = Range("N2", Range("N2").End(xlDown))
'Rango V:AB
    Set l = Range("V2", Range("V2").End(xlDown))
    Set m = Range("W2", Range("W2").End(xlDown))
    Set n = Range("X2", Range("X2").End(xlDown))
    Set o = Range("Y2", Range("Y2").End(xlDown))
    Set p = Range("Z2", Range("Z2").End(xlDown))
    Set q = Range("AA2", Range("AA2").End(xlDown))
    Set r = Range("AB2", Range("AB2").End(xlDown))
'Rango AJ:AO
    Set s = Range("AJ2", Range("AJ2").End(xlDown))
    Set t = Range("AK2", Range("AK2").End(xlDown))
    Set u = Range("AL2", Range("AL2").End(xlDown))
    Set v = Range("AM2", Range("AM2").End(xlDown))
    Set w = Range("AN2", Range("AN2").End(xlDown))
    Set x = Range("AO2", Range("AO2").End(xlDown))
    
    Union(a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x).Select

    For Each celda In Selection
    celda.Value = WorksheetFunction.trim(celda.Value)
    Next

Range("O2", Range("O2").End(xlDown)).NumberFormat = "m/d/yyyy"
Range("AC2", Range("AC2").End(xlDown)).NumberFormat = "m/d/yyyy"

'Transforma el formato de fecha a la que se ocupa en postgres
    Columns("O:O").Select
    Selection.TextToColumns Destination:=Range("O1"), DataType:=xlFixedWidth, _
        fieldinfo:=Array(Array(0, 4), Array(10, 1)), TrailingMinusNumbers:=True
        
    Columns("AC:AC").Select
    Selection.TextToColumns Destination:=Range("AC1"), DataType:=xlFixedWidth, _
        fieldinfo:=Array(Array(0, 4), Array(10, 1)), TrailingMinusNumbers:=True

'Esta instrucción elimina los objetos antes creados para liberar memoria
    Set aa = Nothing
    Set bb = Nothing
    Set cc = Nothing
    Set dd = Nothing
    Set ee = Nothing
    Set a = Nothing
    Set b = Nothing
    Set c = Nothing
    Set d = Nothing
    Set e = Nothing
    Set f = Nothing
    Set g = Nothing
    Set h = Nothing
    Set i = Nothing
    Set j = Nothing
    Set k = Nothing
    Set l = Nothing
    Set m = Nothing
    Set n = Nothing
    Set o = Nothing
    Set p = Nothing
    Set q = Nothing
    Set r = Nothing
    Set s = Nothing
    Set t = Nothing
    Set u = Nothing
    Set v = Nothing
    Set w = Nothing
    Set x = Nothing
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
  
Range("A2").Select

End Sub

Option Explicit

Sub cartoStrack()

'Limpieza de Strack
'Esta acción elimina los bordes de las celdas
'Tambien elimina saltos de página, acentos y diéresis

Dim Rng As Range
Dim WorkRng As Range
Dim celda As Range

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
    .Replace What:="" & Chr(10) & "", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="Á", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="É", Replacement:="E", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="Í", Replacement:="I", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="Ó", Replacement:="O", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="Ú", Replacement:="U", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="Ü", Replacement:="U", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="  ,  , ", Replacement:=", ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" / / / / / ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" / / / / ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" / / / ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" / / ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ///// ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" //// ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" /// ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" // ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'Revisar esta parte del código, si se va a dejar para todas las columnas o se va a trabajar
'Únicamente con las columnas "F:AO"
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

'Convierte en mayúsculas todo el rango en uso
    
    On Error Resume Next
    Selection.CurrentRegion.Select
    Set WorkRng = Application.Selection
    For Each Rng In WorkRng
        Rng.Value = VBA.UCase(Rng.Value)
    Next
    

'Se eliminan los guiones medios del rango AD y AE:AO

    Set a = Range("AD2", Range("AD2").End(xlDown))
    Set b = Range("AE2", Range("AE2").End(xlDown))
    Set c = Range("AF2", Range("AF2").End(xlDown))
    Set d = Range("AH2", Range("AH2").End(xlDown))
    Set e = Range("AI2", Range("AI2").End(xlDown))
    Set f = Range("AJ2", Range("AJ2").End(xlDown))
    Set g = Range("AK2", Range("AK2").End(xlDown))
    Set h = Range("AL2", Range("AL2").End(xlDown))
    Set i = Range("AM2", Range("AM2").End(xlDown))
    Set j = Range("AN2", Range("AN2").End(xlDown))
    Set k = Range("AO2", Range("AO2").End(xlDown))
    
    Union(a, b, c, d, e, f, g, h, i, j, k).Select
    
    With Selection
    
    .Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    End With

'Quita los espacios entre las diagonales dentro del campo TIPO DE DELITO

    Range("D2", Range("D1048576").End(xlUp)).Select
    Selection.Replace What:=" / ", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A2").Select


    
End Sub

Sub gposTacticos()

    'CALIMAYA
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E4:E46") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B4:B46") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3:E46")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CHALCO
    Range("A49").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E50:E179") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B50:B179") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A49:E179")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CHAPULTEPEC
    Range("A182").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E183:E203") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B183:B203") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A182:E203")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CHIMALHUACAN
    Range("A206").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E207:E322") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B207:B322") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A206:E322")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CUAUTITLAN IZCALLI
    Range("A325").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E326:E520") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B326:B520") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A325:E520")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'ECATEPEC
    Range("A523").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E524:E1003") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B524:B1003") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A523:E1003")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'LA PAZ
    Range("A1006").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1007:E1086") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1007:B1086") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1006:E1086")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'LERMA
    Range("A1089").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1090:E1177") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1090:B1177") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1089:E1177")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'METEPEC
    Range("A1180").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1181:E1633") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1181:B1633") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1180:E1633")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'NAUCALPAN
    Range("A1636").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1637:E1991") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1637:B1991") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1636:E1991")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'NEZAHUALCOYOTL
    Range("A1994").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1995:E2081") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1995:B2081") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1994:E2081")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'OTUMBA
    Range("A2084").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2085:E2136") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2085:B2136") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2084:E2136")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'OTZOLOTEPEC
    Range("A2139").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2140:E2183") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2140:B2183") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2139:E2183")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'SAN MATEO ATENCO
    Range("A2186").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2187:E2252") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2187:B2252") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2186:E2252")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TEXCOCO
    Range("A2255").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2256:E2499") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2256:B2499") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2255:E2499")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TOLUCA
    Range("A2502").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2503:E2995") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2503:B2995") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2502:E2995")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TULTITLAN
    Range("A2998").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2999:E3188") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2999:B3188") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2998:E3188")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'VALLE DE CHALCO
    Range("A3191").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3192:E3229") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3192:B3229") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3191:E3229")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'XONACATLAN
    Range("A3232").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3233:E3278") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3233:B3278") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3232:E3278")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'ZINACANTEPEC
    Range("A3281").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3282:E3367") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3282:B3367") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3281:E3367")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("B2").Select
End Sub

Sub gposTacticos()
'Macro para ordenar los datos de los veinte municipios propuestos por el área de Grupos Tácticos
'Esta macro ordena de mayor a menor el conteo de la incidencia por colonia en cada municipio
'De estos datos se obtienen una serie de relojes aorísticos o matrices de la densidad delictiva

Application.ScreenUpdating = False

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
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E50:E178") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B50:B178") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A49:E178")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CHAPULTEPEC
    Range("A181").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E182:E201") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B182:B201") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A181:E201")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CHIMALHUACAN
    Range("A204").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E205:E320") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B205:B320") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A204:E320")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'CUAUTITLAN IZCALLI
    Range("A323").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E324:E518") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B324:B518") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A323:E518")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'ECATEPEC
    Range("A521").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E522:E1001") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B522:B1001") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A521:E1001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'LA PAZ
    Range("A1004").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1005:E1084") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1005:B1084") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1004:E1084")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'LERMA
    Range("A1087").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1088:E1175") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1088:B1175") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1087:E1175")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'METEPEC
    Range("A1178").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1179:E1631") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1179:B1631") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1178:E1631")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'NAUCALPAN
    Range("A1634").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1635:E1990") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1635:B1990") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1634:E1990")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'NEZAHUALCOYOTL
    Range("A1993").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E1994:E2082") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B1994:B2082") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A1993:E2082")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'OTUMBA
    Range("A2085").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2086:E2137") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2086:B2137") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2085:E2137")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'OTZOLOTEPEC
    Range("A2140").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2141:E2184") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2141:B2184") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2140:E2184")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'SAN MATEO ATENCO
    Range("A2187").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2188:E2254") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2188:B2254") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2187:E2254")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TEXCOCO
    Range("A2257").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2258:E2501") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2258:B2501") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2257:E2501")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TOLUCA
    Range("A2504").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E2505:E2996") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B2505:B2996") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2504:E2996")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'TULTITLAN
    Range("A2999").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3000:E3189") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3000:B3189") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A2999:E3189")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'VALLE DE CHALCO
    Range("A3192").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3193:E3230") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3193:B3230") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3192:E3230")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'XONACATLAN
    Range("A3233").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3234:E3279") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3234:B3279") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3233:E3279")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'ZINACANTEPEC
    Range("A3282").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("E3283:E3368") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Rangos").Sort.SortFields.Add2 Key:=Range("B3283:B3368") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Rangos").Sort
        .SetRange Range("A3282:E3368")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("B2").Select
    
Application.ScreenUpdating = True
End Sub

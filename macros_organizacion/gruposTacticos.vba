Sub gposTacticos()
'Macro para ordenar los datos de los 25 municipios propuestos por el área de Grupos Tácticos
'Esta macro ordena de mayor a menor el conteo de la incidencia por colonia en cada municipio
'De estos datos se obtienen una serie de relojes aorísticos o matrices de la densidad delictiva

' https://trumpexcel.com/sort-data-vba/

Application.ScreenUpdating = False

'CALIMAYA
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B4"), Order:=xlAscending 'Colocar el inicio de la serie de datos
     .SortFields.Add Key:=Range("E4"), Order:=xlAscending
     .SetRange Range("A3:E46") 'Se coloca el rango incluyendo los encabezados
     .Header = xlYes
     .Apply
End With

'----------------------------------- CHALCO
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B50"), Order:=xlAscending
     .SortFields.Add Key:=Range("E50"), Order:=xlAscending
     .SetRange Range("A49:E178")
     .Header = xlYes
     .Apply
End With

'----------------------------------- CHAPULTEPEC
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B182"), Order:=xlAscending
     .SortFields.Add Key:=Range("E182"), Order:=xlAscending
     .SetRange Range("A181:E201")
     .Header = xlYes
     .Apply
End With

'----------------------------------- CHIMALHUACAN
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B205"), Order:=xlAscending
     .SortFields.Add Key:=Range("E205"), Order:=xlAscending
     .SetRange Range("A204:E320")
     .Header = xlYes
     .Apply
End With

'----------------------------------- CUAUTITLAN IZCALLI
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B324"), Order:=xlAscending
     .SortFields.Add Key:=Range("E324"), Order:=xlAscending
     .SetRange Range("A323:E518")
     .Header = xlYes
     .Apply
End With


'----------------------------------- ECATEPEC DE MORELOS
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B522"), Order:=xlAscending
     .SortFields.Add Key:=Range("E522"), Order:=xlAscending
     .SetRange Range("A521:E1001")
     .Header = xlYes
     .Apply
End With

'----------------------------------- LA PAZ
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1005"), Order:=xlAscending
     .SortFields.Add Key:=Range("E1005"), Order:=xlAscending
     .SetRange Range("A1004:E1084")
     .Header = xlYes
     .Apply
End With

'----------------------------------- LERMA
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1088"), Order:=xlAscending
     .SortFields.Add Key:=Range("E1088"), Order:=xlAscending
     .SetRange Range("A1087:E1175")
     .Header = xlYes
     .Apply
End With

'----------------------------------- METEPEC
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1179"), Order:=xlAscending
     .SortFields.Add Key:=Range("E1179"), Order:=xlAscending
     .SetRange Range("A1178:E1631")
     .Header = xlYes
     .Apply
End With

'----------------------------------- NAUCALPAN DE JUAREZ
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1635"), Order:=xlAscending
     .SortFields.Add Key:=Range("E1635"), Order:=xlAscending
     .SetRange Range("A1634:E1990")
     .Header = xlYes
     .Apply
End With

'----------------------------------- NEZAHUALCOYOTL
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1994"), Order:=xlAscending
     .SortFields.Add Key:=Range("E1994"), Order:=xlAscending
     .SetRange Range("A1993:E2082")
     .Header = xlYes
     .Apply
End With

'----------------------------------- OTUMBA
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B2086"), Order:=xlAscending
     .SortFields.Add Key:=Range("E2086"), Order:=xlAscending
     .SetRange Range("A2085:E2137")
     .Header = xlYes
     .Apply
End With

'----------------------------------- OTZOLOTEPEC
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B2141"), Order:=xlAscending
     .SortFields.Add Key:=Range("E2141"), Order:=xlAscending
     .SetRange Range("A2140:E2184")
     .Header = xlYes
     .Apply
End With

'----------------------------------- SAN MATEO ATENCO
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B2188"), Order:=xlAscending
     .SortFields.Add Key:=Range("E2188"), Order:=xlAscending
     .SetRange Range("A2187:E2254")
     .Header = xlYes
     .Apply
End With

'----------------------------------- TEXCOCO
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B2258"), Order:=xlAscending
     .SortFields.Add Key:=Range("E2258"), Order:=xlAscending
     .SetRange Range("A2257:E2501")
     .Header = xlYes
     .Apply
End With

'----------------------------------- TOLUCA
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B2505"), Order:=xlAscending
     .SortFields.Add Key:=Range("E2505"), Order:=xlAscending
     .SetRange Range("A2504:E2996")
     .Header = xlYes
     .Apply
End With

'----------------------------------- TULTITLAN
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3000"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3000"), Order:=xlAscending
     .SetRange Range("A2999:E3189")
     .Header = xlYes
     .Apply
End With

'----------------------------------- VALLE DE CHALCO SOLIDARIDAD
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3193"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3193"), Order:=xlAscending
     .SetRange Range("A3192:E3230")
     .Header = xlYes
     .Apply
End With


'----------------------------------- XONACATLAN
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3234"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3234"), Order:=xlAscending
     .SetRange Range("A3233:E3279")
     .Header = xlYes
     .Apply
End With

'----------------------------------- ZINACANTEPEC
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3283"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3283"), Order:=xlAscending
     .SetRange Range("A3282:E3371")
     .Header = xlYes
     .Apply
End With

'----------------------------------- MEXICALTZINGO
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3375"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3375"), Order:=xlAscending
     .SetRange Range("A3374:E3396")
     .Header = xlYes
     .Apply
End With

'----------------------------------- MALINALCO
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3400"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3400"), Order:=xlAscending
     .SetRange Range("A3398:E3443")
     .Header = xlYes
     .Apply
End With

'----------------------------------- OCUILAN
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3447"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3447"), Order:=xlAscending
     .SetRange Range("A3446:E3495")
     .Header = xlYes
     .Apply
End With

'----------------------------------- HUEHUETOCA
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3499"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3499"), Order:=xlAscending
     .SetRange Range("A3498:E3540")
     .Header = xlYes
     .Apply
End With

'----------------------------------- TEPOTZOTLAN
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B3544"), Order:=xlAscending
     .SortFields.Add Key:=Range("E3544"), Order:=xlAscending
     .SetRange Range("A3543:E3590")
     .Header = xlYes
     .Apply
End With

Application.ScreenUpdating = True

End Sub

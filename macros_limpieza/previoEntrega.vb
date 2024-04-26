Sub quitarEspacios()
' Esta macro es para quitar las comas de las columnas L y M que contienen los datos
' de CORPOR y COD_CIERR, puesto que tienen espacios entre las comas que hacen que
' la cadena de texto sea más larga de lo que el campo permite.

' Definimos las variables a utilizar
Dim a, b As Range

ActiveWorkbook.Sheets("ACCIDENTES").Activate

' Definimos los rangos que van a ser utilizados adelante
    Set a = Range("L1", Range("L1048576").End(xlUp))
    Set b = Range("M1", Range("M1048576").End(xlUp))
    
    Union(a, b).Select
    
    Application.ScreenUpdating = False

' Aquí se remplaza el espacio
    With Selection
    
    .Replace What:=" ,", Replacement:=",", LookAt:=xlPart, SearchOrder:=xlByColumns
    
    End With
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

' Se libera la memoria
    Set a = Nothing
    Set b = Nothing

End Sub

Sub concatenarFolio()
'Esta macro concatena la letra A en el folio

    Application.ScreenUpdating = False

' Se elige la columna objetivo
    Columns("C:C").Insert Shift:=xlToRight
    Range("C2").Select
' Se aplica la fórmula de concatenar
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-1],""-A"")"
' Se rellena la columna con los datos a concatenar este paso lo hace hasta el último registro de la columna
    Range("C2").AutoFill Range("C2:C" & Range("B" & Rows.Count).End(xlUp).Row)
' Se copia la información de la columna C
    Range("C2", Range("C2").End(xlDown)).Select
    Selection.Copy
' Se pegan los valores sin formato sobre los datos de la columna B
    Range("B2").PasteSpecial Paste:=xlPasteValues
' Se elimina la columna C, solamente sirvió como base para hacer la operación
    Columns("C:C").Delete Shift:=xlToLeft
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    
End Sub

Sub conocerLength()
'https://foro.todoexcel.com/threads/rellenar-formula-hacia-abajo-hasta-ultima-celda-contigua-con-datos.52932/
    
    Application.ScreenUpdating = False
    
    Columns("M:M").Insert Shift:=xlToRight
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[-1])"
    Range("M2").AutoFill Range("M2:M" & Range("L" & Rows.Count).End(xlUp).Row)

    Columns("O:O").Insert Shift:=xlToRight
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[-1])"
    Range("O2").AutoFill Range("O2:O" & Range("N" & Rows.Count).End(xlUp).Row)
    
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

End Sub
    
Sub previoEntrega()

    quitarEspacios
    concatenarFolio
    conocerLength

End Sub
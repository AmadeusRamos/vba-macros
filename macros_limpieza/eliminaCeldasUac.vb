Sub eliminaCeldasUac()
         
Dim rng As Range
Dim i As Integer, counter As Integer

'Esta macro limpia el rango que se entrega para la UAC al finalizar quedan solamente las claves que se requieren.

Application.ScreenUpdating = False

'Establecer el rango a evaluar como rng.
Set rng = Range("C2", Range("C2").End(xlDown))

'Inicializar i como 1
i = 1

'Loop que realiza el conteo de a 1 sobre las celdas
'en el rango que se desea evaluar
For counter = 1 To rng.Rows.Count

    'Si la celda i en el rango contiene una "x", borrar la celda
    'De lo contrario incrementar i
    If rng.Cells(i) = "10101" Or rng.Cells(i) = "10102" Or rng.Cells(i) = "10103" _
    Or rng.Cells(i) = "10104" Or rng.Cells(i) = "10105" Or rng.Cells(i) = "10106" _
    Or rng.Cells(i) = "10107" Or rng.Cells(i) = "10108" Or rng.Cells(i) = "10111" _
    Or rng.Cells(i) = "10112" Or rng.Cells(i) = "10113" Or rng.Cells(i) = "10114" _
    Or rng.Cells(i) = "10115" Or rng.Cells(i) = "10117" Or rng.Cells(i) = "10118" _
    Or rng.Cells(i) = "10119" Or rng.Cells(i) = "10120" Or rng.Cells(i) = "10121" Then
        rng.Cells(i).EntireRow.Delete
    
            
    Else
        i = i + 1
    End If

Next

'Numeración automática de 1 a n en la columna "A"
For m = 2 To Cells(Rows.Count, "B").End(xlUp).Row
        If Cells(m, "B").Value <> "" Then
         Cells(m, "A").Value = m - 1
        End If
    Next m
    
'Limpia la columna donde está la clave de analista
Range("BA2", Range("BA1048576").End(xlUp)).Clear

Application.ScreenUpdating = True

Range("A2").Select

End Sub
Sub ejercicioCopiar()

    FileCopy "C:\wd\directory\experimentos\macrosExcel\ABRIL 07 08 09 DAI_II.xlsx", "C:\Test\accidentes\ABRIL 07 08 09 DAI_II.xlsx"

End Sub


Sub eliminarFichero()

    Kill "C:\Test\accidentes\ABRIL 07 08 09 DAI_II.xlsx"

End Sub

Sub copiar()

Dim archivo As String

archivo = InputBox("Ingresa nombre de archivo")



End Sub

Sub mayorMenor()

    Dim Edad As Integer
    
    Edad = InputBox("Por favor ingrese una edad", "Ingrese edad")
    
    If Edad >= 18 Then
        MsgBox "Usted es mayor de edad"
    
    Else
        MsgBox "Usted es menor de edad"
    End If

End Sub

Sub ganoExamen()

    Dim Nota As Double
    
    Nota = InputBox("Por favor ingrese la nota")
    
    If Nota >= 3 Then
        MsgBox "Ganó examen"
    Else
    
        MsgBox "Perdió el examen"
    End If
    

End Sub

Sub numeroMayor()

    Dim Numero1 As Double, Numero2 As Double
    
    Numero1 = InputBox("Ingrese un número", "Ingrese número")
    Numero2 = InputBox("Ingrese el segundo número", "Ingrese número")
    
    If Numero1 > Numero2 Then
        MsgBox Numero1
    
    Else
        MsgBox Numero2
    End If


End Sub

'########## CICLO FOR ##########

'Imprimir los números del 1 al 10

Sub Imprimir1a10()

    Dim CadenaNumeros As String
    
    For i = 1 To 10
    
        CadenaNumeros = CadenaNumeros & i & " "
        
    Next i
    MsgBox CadenaNumeros

End Sub

'Imprimir los números del 10 al 1

Sub Imprimir10a1()

    For i = 10 To 1 Step -1
    
        MsgBox i

    Next i
    

End Sub

'########## CICLO Do While ########

'Imprimir los números del 1 al 10

Sub imprimirHasta10()

    Dim contador As Integer
    
    contador = 1
    
    Do While contador <= 10
    
        Debug.Print contador
        contador = contador + 1
        
    Loop

End Sub

'Capturar n números, imprimir si son negativos, positivos o igual a cero
'Cuando el usuario digite la letra x el procedimiento se debe detener y notificar al usuario.

Sub validarNumeros()

    Dim Numero As String
    
    Numero = InputBox("Ingrese un número cualquiera")
    
    Do While Numero <> "x"
    
        If CDbl(Numero) = 0 Then
            MsgBox "El número digitado es igual a cero"
        Else
            If CDbl(Numero) > 0 Then
                MsgBox "El número digitado es positivo"
            Else
                MsgBox "El número digitado es negativo"
            End If
        End If
        Numero = InputBox("Ingrese un nuevo número")
    Loop
    
    MsgBox "El usuario digitó la X y se finalizó el procedimiento"

End Sub

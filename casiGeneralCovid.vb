Sub CasiGeneralCovid()

' Esta Macro realiza la limpieza de símbolos raros.
' Dentro de la columna DAI inserta el texto COVID siempre.
' También agrega SIN DATO y 9999 a un rango de columnas definido.

    Dim dai As String, form_Rv As String, marc_V As String
    Dim submar_V As String, color_V As String, placa_V As String
    Dim laminav_ As String, notav_ As String, arma As String
    Dim hom_Tot As String, hom_Hombr As String, hom_Muj As String, hom_Desc As String
    Dim ultFila As Long
    Dim cont As Long
    
    ultFila = Range("A" & Rows.Count).End(xlUp).Row
    
    For cont = 2 To ultFila
        dai = Cells(cont, 5)
        
        If dai = "" Then
         Cells(cont, 5) = "COVID"
         
        ElseIf dai = "0" Then
         Cells(cont, 5) = "COVID"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        form_Rv = Cells(cont, 34)
        
        If form_Rv = "" Then
         Cells(cont, 34) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        marc_V = Cells(cont, 36)
        
        If marc_V = "" Then
         Cells(cont, 36) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        submar_V = Cells(cont, 37)
        
        If submar_V = "" Then
         Cells(cont, 37) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        color_V = Cells(cont, 38)
        
        If color_V = "" Then
         Cells(cont, 38) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        placa_V = Cells(cont, 39)
        
        If placa_V = "" Then
         Cells(cont, 39) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        laminav_ = Cells(cont, 40)
        
        If laminav_ = "" Then
         Cells(cont, 40) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        notav_ = Cells(cont, 41)
        
        If notav_ = "" Then
         Cells(cont, 41) = "SIN DATO"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        arma = Cells(cont, 46)
        
        If arma = "" Then
         Cells(cont, 46) = "SIN DATO"
         
        End If
    Next cont
                                    
    For cont = 2 To ultFila
        hom_Tot = Cells(cont, 42)
        
        If hom_Tot = "" Then
         Cells(cont, 42) = "9999"
         
        End If
    Next cont
    
    For cont = 2 To ultFila
        hom_Hombr = Cells(cont, 43)
        
        If hom_Hombr = "" Then
         Cells(cont, 43) = "9999"
         
        End If
    Next cont
        
    For cont = 2 To ultFila
        hom_Muj = Cells(cont, 44)
        
        If hom_Muj = "" Then
         Cells(cont, 44) = "9999"
         
        End If
    Next cont
            
    For cont = 2 To ultFila
        hom_Desc = Cells(cont, 45)
        
        If hom_Desc = "" Then
         Cells(cont, 45) = "9999"
         
        End If
    Next cont
                                
        

Columns("F:AB").Replace _
 What:="•", Replacement:="", _
 SearchOrder:=xlByRows, MatchCase:=True
    Columns("F:AB").Replace _
 What:="$", Replacement:="", _
 SearchOrder:=xlByRows, MatchCase:=True
    Columns("F:AB").Replace _
 What:="|", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="°", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="!", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="¡", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="""", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="#", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="%", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="&", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="(", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:=")", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
     Columns("F:AB").Replace _
 What:="=", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="~*", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 Columns("F:AB").Replace _
 What:="~?", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="~¿", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="'", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="+", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="{", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="}", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="[", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="]", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
     Columns("F:AB").Replace _
 What:="<", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:=">", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="-", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="`", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
    Columns("F:AB").Replace _
 What:="" & Chr(10) & "", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

End Sub
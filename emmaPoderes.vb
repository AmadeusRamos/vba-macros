Option Explicit

Sub emmaPoderes()

Application.ScreenUpdating = False

Dim corpor As String, cod_Cierr As String, hms As String, call_Inc As String
Dim zona As String, reg_Op As String, modo_Rec As String, dir_Inc As String
Dim entre_Inc As String, col_Inc As String, ref_Inc As String, not_Inc As String
Dim fec_Rv As String, aaaa_Rv As String, mm_Rv As String, dd_Rv As String
Dim h_Rv As String, form_Rv As String, mod_V As String, marc_V As String
Dim submar_V As String, color_V As String, placa_V As String, laminav_ As String
Dim notav_ As String, hom_Tot As String, hom_Hombr As String, hom_Muj As String
Dim hom_Desc As String, arma As String
Dim ultFila As Long
Dim cont As Long
Dim celda As Range

'>>>>>>>>>>>>>>>>>>>>PRIMERA PARTE<<<<<<<<<<<<<<<<<<<<

'Limpia las columnas donde existe información después de una coma
'Dejando solamente el primer dato que es útil para la base de datos
'Que se actualiza en Postgres constantemente

'Columna MOD_V
    
    Columns("AJ:AS").Insert Shift:=xlToRight
    Range("AI2", Range("AI1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AI2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AJ:AS").Select
    Selection.Delete Shift:=xlToLeft
    
'Columna MARC_V

    Columns("AK:AT").Insert Shift:=xlToRight
    Range("AJ2", Range("AJ1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AJ2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AK:AT").Select
    Selection.Delete Shift:=xlToLeft
    
'Columna SUBMAR_V

    Columns("AL:AU").Insert Shift:=xlToRight
    Range("AK2", Range("AK1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AK2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AL:AU").Select
    Selection.Delete Shift:=xlToLeft
        
'Columna COLOR_V
    
    Columns("AM:AV").Insert Shift:=xlToRight
    Range("AL2", Range("AL1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AL2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AM:AV").Select
    Selection.Delete Shift:=xlToLeft
        
'Columna PLACA_V

    Columns("AN:AW").Insert Shift:=xlToRight
    Range("AM2", Range("AM1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AM2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AN:AW").Select
    Selection.Delete Shift:=xlToLeft
    
'Columna LAMINAV_

    Columns("AO:AX").Insert Shift:=xlToRight
    Range("AN2", Range("AN1048576").End(xlUp)).Select
    Selection.TextToColumns Destination:=Range("AN2"), DataType:=xlDelimited, _
        textqualifier:=xlDoubleQuote, Tab:=True, comma:=True, fieldinfo:=Array(1, 1)
    Columns("AO:AX").Select
    Selection.Delete Shift:=xlToLeft

'>>>>>>>>>>>>>>>>>>>>SEGUNDA PARTE<<<<<<<<<<<<<<<<<<<<

'Este segmento extrae del campo FECHA_RV los datos y los va agregando en
'Las tres columnas siguientes AAAA_RV, MM_RV, DD_RV

    Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AD2")
    Range("AD2", Range("AD2").End(xlDown)).Select
    
    With Selection
    
    .FormulaR1C1 = "=TEXT(RC[-1],""aaaa"")"
    .Copy
    .PasteSpecial Paste:=xlPasteValues
    .NumberFormat = "@"
    
    End With
    
    Columns("AE:AE").Select
    Selection.Insert Shift:=xlToRight
    Range("AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AE2")
    Range("AE2", Range("AE2").End(xlDown)).Select
    Selection.NumberFormat = "General"
    
    With Selection
    
    .FormulaR1C1 = "=NUMBERVALUE(RC[-1])"
    .Copy
    
    End With
    
    Range("AD2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.NumberFormat = "General"
    Columns("AE:AE").Select
    Selection.Delete Shift:=xlToLeft
    Range("AD2").Select
    
'Esta parte cambia el formato de destino de fecha al mes en curso

    Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AE2")
    Range("AE2", Range("AE2").End(xlDown)).Select
    
    With Selection
    
    .FormulaR1C1 = "=TEXT(RC[-2],""mmmm"")"
    .Copy
    .PasteSpecial Paste:=xlPasteValues
    .NumberFormat = "@"
    
    End With
    
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight
    Range("AE2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AF2")
    Range("AF2", Range("AF2").End(xlDown)).Select
    Selection.NumberFormat = "General"
    
    With Selection
    
    .FormulaR1C1 = "=UPPER(RC[-1])"
    .Copy
    
    End With
    
    Range("AE2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Columns("AF:AF").Select
    Selection.Delete Shift:=xlToLeft
    Range("AF2").Select
    
'Cambio de día
    
    Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AF2")
    Range("AF2", Range("AF2").End(xlDown)).Select
    
    With Selection
    
    .FormulaR1C1 = "=TEXT(RC[-3],""dddd"")"
    .Copy
    .PasteSpecial Paste:=xlPasteValues
    .NumberFormat = "@"
    
    End With
    
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight
    Range("AF2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Destination:=Range("AG2")
    Range("AG2", Range("AG2").End(xlDown)).Select
    Selection.NumberFormat = "General"
    
    With Selection
    
    .FormulaR1C1 = "=UPPER(SUBSTITUTE(SUBSTITUTE(TEXT(RC[-1],""dddd""),""á"",""a""),""é"",""e""))"
    .Copy
    
    End With
    
    Range("AF2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Columns("AG:AG").Select
    Selection.Delete Shift:=xlToLeft
    Range("AF2").Select
    
' Esta sección cambia el formato de hh:mm a hh:mm:ss
' Comienza con la columna "T"

    Columns("U:U").Insert Shift:=xlToRight
    Range("T2").Select
    Range(Selection, Selection.End(xlDown)).Copy Destination:=Range("U2")
    Range("U2", Range("U2").End(xlDown)).Select
               
    With Selection

        .FormulaR1C1 = "=TEXT(RC[-1],""hh:mm:ss"")"
        .Copy
        .PasteSpecial Paste:=xlPasteValues
        .NumberFormat = "hh:mm:ss"
        .Copy
                
    End With
    
    Range("T2").PasteSpecial Paste:=xlPasteValues
    Columns("U:U").Delete Shift:=xlToLeft

'Continúa con la columna "AG" donde está el siguiente dato

    Columns("AH:AH").Insert Shift:=xlToRight
    Range("AG2").Select
    Range(Selection, Selection.End(xlDown)).Copy Destination:=Range("AH2")
    Range("AH2", Range("AH2").End(xlDown)).Select
    
    With Selection
    
        .FormulaR1C1 = "=TEXT(RC[-1], ""hh:mm:ss"")"
        .Copy
        .PasteSpecial Paste:=xlPasteValues
        .NumberFormat = "hh:mm:ss"
        .Copy
        
    End With
    
    Range("AG2").PasteSpecial Paste:=xlPasteValues
    Columns("AH:AH").Delete Shift:=xlToLeft
    Range("AG2").Select


'>>>>>>>>>>>>>>>>>>>>TERCERA PARTE<<<<<<<<<<<<<<<<<<<<

'En esta sección se van a quitar texto y números no deseados de los rangos
'"J:N", "W:AB", "AG" y de "AI:AO"

ultFila = Range("A" & rows.Count).End(xlUp).Row
    
'ZONA
    
    For cont = 2 To ultFila
        zona = Cells(cont, 10)
        
        If zona = "" Then
         Cells(cont, 10) = "SIN DATO"
         
        End If
    Next cont

'REG_OP

    For cont = 2 To ultFila
        reg_Op = Cells(cont, 11)
        
        If reg_Op = "" Then
         Cells(cont, 11) = "SIN DATO"
        
        End If
    Next cont

'CORPOR

    For cont = 2 To ultFila
        corpor = Cells(cont, 12)
        
        If corpor = "" Or corpor = "0" Or corpor = "-" Or corpor = " - " Then
         Cells(cont, 12) = "SIN DATO"
         
        End If
    Next cont
    
'COD_CIERR

    For cont = 2 To ultFila
        cod_Cierr = Cells(cont, 13)
        
        If cod_Cierr = "" Or cod_Cierr = "0" Or cod_Cierr = "-" Or cod_Cierr = " - " Then
         Cells(cont, 13) = "SIN DATO"
         
        End If
    Next cont
    
'MODO_REC

    For cont = 2 To ultFila
        modo_Rec = Cells(cont, 14)
        
        If modo_Rec = "" Then
         Cells(cont, 14) = "SIN DATO"
         
        End If
    Next cont

'DIR_INC

    For cont = 2 To ultFila
        dir_Inc = Cells(cont, 23)
        
        If dir_Inc = "" Then
         Cells(cont, 23) = "SIN DATO"
        
        End If
    Next cont

'CALL_INC

    For cont = 2 To ultFila
        call_Inc = Cells(cont, 24)
        
        If call_Inc = "" Or call_Inc = "0" Or call_Inc = "-" Or call_Inc = " - " Or call_Inc = "M" Or call_Inc = "N" Or call_Inc = "N.P" Or call_Inc = "N.P." Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "N/P" Or call_Inc = "N´P" Or call_Inc = "NA" Or call_Inc = "ND" Or call_Inc = "NINGUNA" Or call_Inc = "NINGUNO" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "NNINGUNA" Or call_Inc = "NNINGUNO" Or call_Inc = "NO" Or call_Inc = "NO  P." Or call_Inc = "NO INDICA" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "NO P." Or call_Inc = "NO PROPORCIONA" Or call_Inc = "NO PROPORCIONA MAS INFORMACION" Or call_Inc = "NOINDICA" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "NOPROPORCIONA" Or call_Inc = "NP" Or call_Inc = "N-P" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "NP." Or call_Inc = "NPN" Or call_Inc = "NPNP" Or call_Inc = "NPNPNP" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "OTRA" Or call_Inc = "OTRAS" Or call_Inc = "P" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "PNP" Or call_Inc = "S" Or call_Inc = "S.D" Or call_Inc = "S.D." Or call_Inc = "S.N" Or call_Inc = "S.N." Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "S.P" Or call_Inc = "S.P." Or call_Inc = "S/D" Or call_Inc = "SA" Or call_Inc = "SC" Or call_Inc = "SD" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SD." Or call_Inc = "SIN" Or call_Inc = "SIN C" Or call_Inc = "SIN CALE" Or call_Inc = "SIN CALL E" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SIN CALLE" Or call_Inc = "SIN CALLES" Or call_Inc = "SIN CLLE" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SIN DATOS" Or call_Inc = "SIN ESPECIFICAR" Or call_Inc = "SIN INFORMACION" Or call_Inc = "SIN NOMBRE" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SINCALLE" Or call_Inc = "SINCALLES" Or call_Inc = "SINDATO" Or call_Inc = "SINDATOS" Or call_Inc = "SINESPECIFICAR" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SININFORMACION" Or call_Inc = "SINNOMBRE" Then
         Cells(cont, 24) = "SIN DATO"
         
        ElseIf call_Inc = "SN" Or call_Inc = "SN." Or call_Inc = "SNI CALLE" Or call_Inc = "SP" Or call_Inc = "SP." Or call_Inc = " SIN ESPECIFICAR" Then
         Cells(cont, 24) = "SIN DATO"
                 
        End If
    Next cont

'ENTRE_INC
    
    For cont = 2 To ultFila
        entre_Inc = Cells(cont, 25)
        
        If entre_Inc = "" Or entre_Inc = "0" Or entre_Inc = "-" Or entre_Inc = " - " Or entre_Inc = "M" Or entre_Inc = "N" Or entre_Inc = "N.P" Or entre_Inc = "N.P." Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "N/P" Or entre_Inc = "N´P" Or entre_Inc = "NA" Or entre_Inc = "ND" Or entre_Inc = "NINGUNA" Or entre_Inc = "NINGUNO" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "NNINGUNA" Or entre_Inc = "NNINGUNO" Or entre_Inc = "NO" Or entre_Inc = "NO  P." Or entre_Inc = "NO INDICA" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "NO P." Or entre_Inc = "NO PROPORCIONA" Or entre_Inc = "NO PROPORCIONA MAS INFORMACION" Or entre_Inc = "NOINDICA" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "NOPROPORCIONA" Or entre_Inc = "NP" Or entre_Inc = "N-P" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "NP." Or entre_Inc = "NPN" Or entre_Inc = "NPNP" Or entre_Inc = "NPNPNP" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "OTRA" Or entre_Inc = "OTRAS" Or entre_Inc = "P" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "PNP" Or entre_Inc = "S" Or entre_Inc = "S.D" Or entre_Inc = "S.D." Or entre_Inc = "S.N" Or entre_Inc = "S.N." Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "S.P" Or entre_Inc = "S.P." Or entre_Inc = "S/D" Or entre_Inc = "SA" Or entre_Inc = "SC" Or entre_Inc = "SD" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SD." Or entre_Inc = "SIN" Or entre_Inc = "SIN C" Or entre_Inc = "SIN CALE" Or entre_Inc = "SIN CALL E" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SIN CALLE" Or entre_Inc = "SIN CALLES" Or entre_Inc = "SIN CLLE" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SIN DATOS" Or entre_Inc = "SIN ESPECIFICAR" Or entre_Inc = "SIN INFORMACION" Or entre_Inc = "SIN NOMBRE" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SINCALLE" Or entre_Inc = "SINCALLES" Or entre_Inc = "SINDATO" Or entre_Inc = "SINDATOS" Or entre_Inc = "SINESPECIFICAR" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SININFORMACION" Or entre_Inc = "SINNOMBRE" Then
         Cells(cont, 25) = "SIN DATO"
         
        ElseIf entre_Inc = "SN" Or entre_Inc = "SN." Or entre_Inc = "SNI CALLE" Or entre_Inc = "SP" Or entre_Inc = "SP." Or entre_Inc = " SIN ESPECIFICAR" Then
         Cells(cont, 25) = "SIN DATO"
                 
        End If
    Next cont

'COL_INC
    
    For cont = 2 To ultFila
        col_Inc = Cells(cont, 26)
        
        If col_Inc = "" Or col_Inc = "0" Or col_Inc = "-" Or col_Inc = " - " Or col_Inc = "M" Or col_Inc = "N" Or col_Inc = "N.P" Or col_Inc = "N.P." Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "N/P" Or col_Inc = "N´P" Or col_Inc = "NA" Or col_Inc = "ND" Or col_Inc = "NINGUNA" Or col_Inc = "NINGUNO" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "NNINGUNA" Or col_Inc = "NNINGUNO" Or col_Inc = "NO" Or col_Inc = "NO  P." Or col_Inc = "NO INDICA" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "NO P." Or col_Inc = "NO PROPORCIONA" Or col_Inc = "NO PROPORCIONA MAS INFORMACION" Or col_Inc = "NOINDICA" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "NOPROPORCIONA" Or col_Inc = "NP" Or col_Inc = "N-P" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "NP." Or col_Inc = "NPN" Or col_Inc = "NPNP" Or col_Inc = "NPNPNP" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "OTRA" Or col_Inc = "OTRAS" Or col_Inc = "P" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "PNP" Or col_Inc = "S" Or col_Inc = "S.D" Or col_Inc = "S.D." Or col_Inc = "S.N" Or col_Inc = "S.N." Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "S.P" Or col_Inc = "S.P." Or col_Inc = "S/D" Or col_Inc = "SA" Or col_Inc = "SC" Or col_Inc = "SD" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "SD." Or col_Inc = "SIN" Or col_Inc = "SIN C" Or col_Inc = "SIN COLONIA" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "SIN DATOS" Or col_Inc = "SIN ESPECIFICAR" Or col_Inc = "SIN INFORMACION" Or col_Inc = "SIN NOMBRE" Then
         Cells(cont, 26) = "SIN DATO"
  
        ElseIf col_Inc = "SINDATO" Or col_Inc = "SINDATOS" Or col_Inc = "SINESPECIFICAR" Then
         Cells(cont, 26) = "SIN DATO"
         
        ElseIf col_Inc = "SININFORMACION" Or col_Inc = "SINNOMBRE" Or col_Inc = "SN" Or col_Inc = "SN." Then
         Cells(cont, 26) = "SIN DATO"
    
        ElseIf col_Inc = "SP" Or col_Inc = "SP." Or col_Inc = " SIN ESPECIFICAR" Then
         Cells(cont, 26) = "SIN DATO"
        
        End If
    Next cont

'REF_INC
    
    For cont = 2 To ultFila
        ref_Inc = Cells(cont, 27)
        
        If ref_Inc = "" Or ref_Inc = "0" Or ref_Inc = "-" Or ref_Inc = " - " Or ref_Inc = "M" Or ref_Inc = "N" Or ref_Inc = "N.P" Or ref_Inc = "N.P." Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "N/P" Or ref_Inc = "N´P" Or ref_Inc = "NA" Or ref_Inc = "ND" Or ref_Inc = "NINGUNA" Or ref_Inc = "NINGUNO" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "NNINGUNA" Or ref_Inc = "NNINGUNO" Or ref_Inc = "NO" Or ref_Inc = "NO  P." Or ref_Inc = "NO INDICA" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "NO P." Or ref_Inc = "NO PROPORCIONA" Or ref_Inc = "NO PROPORCIONA MAS INFORMACION" Or ref_Inc = "NOINDICA" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "NOPROPORCIONA" Or ref_Inc = "NP" Or ref_Inc = "N-P" Or ref_Inc = "NP REFERENCIA" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "NP REFERENCIAS" Or ref_Inc = "NP." Or ref_Inc = "NPN" Or ref_Inc = "NPNP" Or ref_Inc = "NPNPNP" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "NPREFERENCIA" Or ref_Inc = "NPREFERENCIAS" Or ref_Inc = "OTRA" Or ref_Inc = "OTRAS" Or ref_Inc = "P" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "PNP" Or ref_Inc = "S" Or ref_Inc = "S.D" Or ref_Inc = "S.D." Or ref_Inc = "S.N" Or ref_Inc = "S.N." Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "S.P" Or ref_Inc = "S.P." Or ref_Inc = "S/D" Or ref_Inc = "SA" Or ref_Inc = "SC" Or ref_Inc = "SD" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SD." Or ref_Inc = "SIN" Or ref_Inc = "SIN C" Or ref_Inc = "SIN CALE" Or ref_Inc = "SIN CALL E" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SIN CALLE" Or ref_Inc = "SIN CALLES" Or ref_Inc = "SIN CLLE" Or ref_Inc = "SIN COLONIA" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SIN DATOS" Or ref_Inc = "SIN ESPECIFICAR" Or ref_Inc = "SIN INFORMACION" Or ref_Inc = "SIN NOMBRE" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SIN PLACA" Or ref_Inc = "SIN PLACAS" Or ref_Inc = "SIN REFERENCIA" Or ref_Inc = "SIN REFERENCIAS " Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SINCALLE" Or ref_Inc = "SINCALLES" Or ref_Inc = "SINDATO" Or ref_Inc = "SINDATOS" Or ref_Inc = "SINESPECIFICAR" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SININFORMACION" Or ref_Inc = "SINNOMBRE" Or ref_Inc = "SINPLAC" Or ref_Inc = "SINPLACA" Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SINPLACAS" Or ref_Inc = "SINREFERENCIA" Or ref_Inc = "SINREFERENCIAS" Or ref_Inc = "SN" Or ref_Inc = "SN." Then
         Cells(cont, 27) = "SIN DATO"
         
        ElseIf ref_Inc = "SNI CALLE" Or ref_Inc = "SP" Or ref_Inc = "SP." Or ref_Inc = "N P" Or ref_Inc = "NNP" Or ref_Inc = " SIN ESPECIFICAR" Then
         Cells(cont, 27) = "SIN DATO"

        ElseIf ref_Inc = "NADA" Or ref_Inc = "NO INDICO" Or ref_Inc = "NP -REFERENCIA" Or ref_Inc = "." Then
         Cells(cont, 27) = "SIN DATO"
        
        End If
    Next cont

'NOT_INC

    For cont = 2 To ultFila
        not_Inc = Cells(cont, 28)
        
        If not_Inc = "" Or not_Inc = "0" Or not_Inc = "-" Or not_Inc = " - " Or not_Inc = "M" Or not_Inc = "N" Or not_Inc = "N.P" Or not_Inc = "N.P." Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "N/P" Or not_Inc = "N´P" Or not_Inc = "NA" Or not_Inc = "ND" Or not_Inc = "NINGUNA" Or not_Inc = "NINGUNO" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "NNINGUNA" Or not_Inc = "NNINGUNO" Or not_Inc = "NO" Or not_Inc = "NO  P." Or not_Inc = "NO INDICA" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "NO P." Or not_Inc = "NO PROPORCIONA" Or not_Inc = "NO PROPORCIONA MAS INFORMACION" Or not_Inc = "NOINDICA" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "NOPROPORCIONA" Or not_Inc = "NP" Or not_Inc = "N-P" Or not_Inc = "NP REFERENCIA" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "NP REFERENCIAS" Or not_Inc = "NP." Or not_Inc = "NPN" Or not_Inc = "NPNP" Or not_Inc = "NPNPNP" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "NPREFERENCIA" Or not_Inc = "NPREFERENCIAS" Or not_Inc = "OTRA" Or not_Inc = "OTRAS" Or not_Inc = "P" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "PNP" Or not_Inc = "S" Or not_Inc = "S.D" Or not_Inc = "S.D." Or not_Inc = "S.N" Or not_Inc = "S.N." Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "S.P" Or not_Inc = "S.P." Or not_Inc = "S/D" Or not_Inc = "SA" Or not_Inc = "SC" Or not_Inc = "SD" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SD." Or not_Inc = "SIN" Or not_Inc = "SIN C" Or not_Inc = "SIN CALE" Or not_Inc = "SIN CALL E" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SIN CALLE" Or not_Inc = "SIN CALLES" Or not_Inc = "SIN CLLE" Or not_Inc = "SIN COLONIA" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SIN DATOS" Or not_Inc = "SIN ESPECIFICAR" Or not_Inc = "SIN INFORMACION" Or not_Inc = "SIN NOMBRE" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SIN PLACA" Or not_Inc = "SIN PLACAS" Or not_Inc = "SIN REFERENCIA" Or not_Inc = "SIN REFERENCIAS " Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SINCALLE" Or not_Inc = "SINCALLES" Or not_Inc = "SINDATO" Or not_Inc = "SINDATOS" Or not_Inc = "SINESPECIFICAR" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SININFORMACION" Or not_Inc = "SINNOMBRE" Or not_Inc = "SINPLAC" Or not_Inc = "SINPLACA" Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SINPLACAS" Or not_Inc = "SINREFERENCIA" Or not_Inc = "SINREFERENCIAS" Or not_Inc = "SN" Or not_Inc = "SN." Then
         Cells(cont, 28) = "SIN DATO"
         
        ElseIf not_Inc = "SNI CALLE" Or not_Inc = "SP" Or not_Inc = "SP." Or not_Inc = " SIN ESPECIFICAR" Then
         Cells(cont, 28) = "SIN DATO"
        
        End If
    Next cont

'H_RV

    For cont = 2 To ultFila
        h_Rv = Cells(cont, 33)
        
        If h_Rv = "" Then
         Cells(cont, 33) = "SIN DATO"
         
        End If
    Next cont

'FORM_RV

    For cont = 2 To ultFila
        form_Rv = Cells(cont, 34)
        
        If form_Rv = "" Or form_Rv = "0" Or form_Rv = "-" Or form_Rv = " - " Or form_Rv = "SIN INFORMACION" Then
         Cells(cont, 34) = "SIN DATO"
        
        End If
    Next cont

'MOD_V
       
    For cont = 2 To ultFila
        mod_V = Cells(cont, 35)
        
        If mod_V = "" Or mod_V = "-" Or mod_V = " - " Then
         Cells(cont, 35) = "0"
         
        End If
    Next cont
    
'MARC_V

    For cont = 2 To ultFila
        marc_V = Cells(cont, 36)
        
        If marc_V = "" Or marc_V = "0" Or marc_V = "-" Or marc_V = " - " Or marc_V = "SIN INFORMACION" Or marc_V = " SIN INFORMACION " Then
         Cells(cont, 36) = "SIN DATO"
         
        End If
    Next cont
    
'SUBMAR_V

    For cont = 2 To ultFila
        submar_V = Cells(cont, 37)
        
        If submar_V = "" Or submar_V = "0" Or submar_V = "-" Or submar_V = " - " Or submar_V = "SIN INFORMACION" Or submar_V = " SIN INFORMACION " Then
         Cells(cont, 37) = "SIN DATO"
         
        End If
    Next cont
    
'COLOR_V

    For cont = 2 To ultFila
        color_V = Cells(cont, 38)
        
        If color_V = "" Or color_V = "0" Or color_V = "-" Or color_V = " - " Or color_V = "SIN INFORMACION" Or color_V = " SIN INFORMACION " Then
         Cells(cont, 38) = "SIN DATO"
         
        End If
    Next cont

'PLACA_V

    For cont = 2 To ultFila
        placa_V = Cells(cont, 39)
        
        If placa_V = "" Or placa_V = " " Or placa_V = "0" Or placa_V = "-" Or placa_V = " - " Or placa_V = "00" Or placa_V = "000" Or placa_V = "0000" Then
         Cells(cont, 39) = "SIN DATO"
        
        ElseIf placa_V = "SD" Or placa_V = "S/D" Or placa_V = "SP" Or placa_V = "NP" Or placa_V = "NA" Or placa_V = "SIN PLACA" Then
         Cells(cont, 39) = "SIN DATO"

        ElseIf placa_V = "SINPLACA" Or placa_V = "SIN NUMERO" Or placa_V = "SINNUM" Or placa_V = "SINNUME" Or placa_V = " SIN INFORMACION " Then
         Cells(cont, 39) = "SIN DATO"

        ElseIf placa_V = " NP " Or placa_V = "SINDATO" Or placa_V = "SIN" Or placa_V = "SINPLAC" Or placa_V = " SIN " Or placa_V = "NT" Then
         Cells(cont, 39) = "SIN DATO"
        
        End If
    Next cont

'LAMINAV_

    For cont = 2 To ultFila
        laminav_ = Cells(cont, 40)
        
        If laminav_ = "" Or laminav_ = "0" Or laminav_ = "-" Or laminav_ = " - " Or laminav_ = "SIN ESPECIFICAR" Or laminav_ = "SINESPECIFICAR" Then
         Cells(cont, 40) = "SIN DATO"
         
        ElseIf laminav_ = "NP" Or laminav_ = "SD" Or laminav_ = "SP" Or laminav_ = "," Or laminav_ = " , " Or laminav_ = " SIN ESPECIFICAR" Or laminav_ = " SIN INFORMACION " Then
         Cells(cont, 40) = "SIN DATO"
         
        End If
    Next cont

'NOTAV_

    For cont = 2 To ultFila
        notav_ = Cells(cont, 41)
        
        If notav_ = "" Or notav_ = "0" Or notav_ = "-" Or notav_ = " - " Or notav_ = "NP" Or notav_ = "NO PROPORCIONA" Or notav_ = " SIN INFORMACION " Then
         Cells(cont, 41) = "SIN DATO"

        ElseIf notav_ = "NINGUNA" Or notav_ = "NINGUNO" Or notav_ = "NINGUN A" Or notav_ = "SIN SEÑAS PARTICULARES" Or notav_ = "SIN SEÑAS" Then
         Cells(cont, 41) = "SIN DATO"
         
        End If
    Next cont

'Esta sección va a colocar en las celdas que están vacías dentro del rango
'"AP:AT" la información SIN DATO y 9999

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

'>>>>>>>>>>>>>>>>>>>>CUARTA PARTE<<<<<<<<<<<<<<<<<<<<
   
'Este proceso cambia la fuente de las celdas.
'Así como su posición y tamaño

    Range("A2").CurrentRegion.Select

    With Selection

    .Font.Name = "Arial"
    .Font.Size = 10
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Orientation = 0
    .IndentLevel = 0
    .ReadingOrder = xlContext

    End With

'Esta línea transforma la columna PLACA_V "AM" a texto
'Para quitar las celdas que están en formato científico

Range("AM2", Range("AM2").End(xlDown)).NumberFormat = "@"
Range("O2", Range("O2").End(xlDown)).NumberFormat = "m/d/yyyy"
Range("AC2", Range("AC2").End(xlDown)).NumberFormat = "m/d/yyyy"

Application.ScreenUpdating = True


    Range("A2").Select

End Sub

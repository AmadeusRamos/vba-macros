Sub limpiarWorkspace()

Application.ScreenUpdating = False

Range("A2", Range("A2").End(xlDown)).Select
Selection.resize(, 7).Select


    With Selection

    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Color = vbBlack
    .Orientation = 0
    .IndentLevel = 0
    .ReadingOrder = xlContext
    .Clear

    End With
    
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    
Application.ScreenUpdating = True
Application.CutCopyMode = False
    
Range("A2").Select

End Sub

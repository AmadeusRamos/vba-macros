Sub eliminarFilas()
         
'Esta macro elimina toda la fila de las celdas seleccionadas
         
Application.ScreenUpdating = False
         
   Selection.EntireRow.Delete
   
Application.ScreenUpdating = True

Range("A2").Select
    
End Sub

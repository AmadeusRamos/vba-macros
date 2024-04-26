'Sub limpiarPortapapeles()

'EvRClearOfficeClipBoard

'End Sub

#If VBA7 Then
    Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
    ByVal iChildStart As Long, ByVal cChildren As Long, _
    ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
    Public Const myVBA7 As Long = 1
'#Else
    'Private Declare Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                              ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                              ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
    'Public Const myVBA7 As Long = 0
#End If

Public Sub EvRClearOfficeClipBoard()
'https://stackoverflow.com/questions/64066265/clearing-the-clipboard-in-office-365

    Dim cmnB, IsVis As Boolean, j As Long, Arr As Variant
    Arr = Array(4, 7, 2, 0)                      '4 and 2 for 32 bit, 7 and 0 for 64 bit
    Set cmnB = Application.CommandBars("Office Clipboard")
    IsVis = cmnB.Visible
    If Not IsVis Then
        cmnB.Visible = True
        DoEvents
    End If

    For j = 1 To Arr(0 + myVBA7)
        AccessibleChildren cmnB, Choose(j, 0, 3, 0, 3, 0, 3, 1), 1, cmnB, 1
    Next
        
    cmnB.accDoDefaultAction CLng(Arr(2 + myVBA7))

    Application.CommandBars("Office Clipboard").Visible = IsVis

End Sub

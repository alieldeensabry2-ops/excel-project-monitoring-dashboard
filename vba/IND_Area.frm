VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IND_Area 
   ClientHeight    =   8760.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12390
   OleObjectBlob   =   "IND_Area.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IND_Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" ( _
    ByVal hCursor As LongPtr) As LongPtr

Const IDC_HAND As Long = 32649 ' —„“ √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…




Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
    
    For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
    
    Me.IN_B1.Visible = False
End Sub


Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B2.Visible = False
End Sub
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B3.Visible = False
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B4.Visible = False
End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B5.Visible = False
End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B6.Visible = False
End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B7.Visible = False
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B8.Visible = False
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B9.Visible = False
End Sub
Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B10.Visible = False
End Sub
Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B11.Visible = False
End Sub
Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
        For i = 1 To 12
        Me.Controls("IN_B" & i).Visible = True
    Next i
        Me.IN_B12.Visible = False
End Sub


'Buttons actions -----------------------
Public Sub Label1_Click()
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A4"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub Label10_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A13"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
End Sub

Public Sub Label11_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A14"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label12_Click()
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A15"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub Label2_Click()
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A5"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub Label3_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A6"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label4_Click()
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A7"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub Label5_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A8"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label6_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
        
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A9"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label7_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A10"
    
        ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label8_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A11"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub Label9_Click()

    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage2", "MultiPage3", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A12"
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub
'------------------------------

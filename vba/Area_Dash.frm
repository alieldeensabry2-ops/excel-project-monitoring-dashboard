VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Area_Dash 
   ClientHeight    =   13335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23955
   OleObjectBlob   =   "Area_Dash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Area_Dash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" ( _
    ByVal hCursor As LongPtr) As LongPtr

Const IDC_HAND As Long = 32649 ' —„“ √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…

Private Sub Image13_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label35_Click()
 Dim ws As Worksheet
    Dim pt As PivotTable
    Dim projectName As String
    Dim companyName As String
    Dim resultRange As Range
    Dim labelIndex As Integer
    Dim cell As Range
    Dim lbl As Object
    Dim img As Object
    Dim imgA As Object
    Dim imgB As Object
    Dim lblc As Object
    Dim visibleLabelsc As Integer
    ' Get the value from cell B2 (for "Project")
    Set ws = ThisWorkbook.Sheets("Projects_Tab_pivots")
    projectName = ws.Range("B2").Value
    
    ' Set the Pivot Table for filtering "Project" and "Company Name"
    Set pt = ws.PivotTables("Contracts_Info")
    
    ' Apply filter to the "Project" field based on value in B2
    pt.PivotFields("Project").ClearAllFilters
    pt.PivotFields("Project").CurrentPage = projectName
    
    ' Apply filter to the "Company Name" field based on the selected label
    pt.PivotFields("Company Name").ClearAllFilters
    
    ' Get the results from the PivotTable in column O
    Set resultRange = ws.Range("O6:O" & ws.Cells(ws.Rows.Count, "O").End(xlUp).Row)
    
    ' Step 1: Hide all the labels and images from LBLC1 to LBLC20 and Imc1 to Imc20
    For i = 1 To 20
        Set lbl = Area_Dash.Controls("LBLC" & i)
        lbl.Visible = False ' Hide the label
        
        Set img = Area_Dash.Controls("Imc" & i)
        img.Visible = False ' Hide the corresponding image
         Set imgA = Area_Dash.Controls("Imd" & i)
        imgA.Visible = False ' Hide the corresponding image
        Set imgB = Area_Dash.Controls("LBLCA" & i)
        imgB.Visible = False ' Hide the corresponding image
    Next i
    
    ' Step 2: Loop through the result range and update the labels
    labelIndex = 1
    For Each cell In resultRange
        If InStr(cell.Value, "Grand Total") > 0 Then
            Exit For ' Exit the loop if "Grand Total" is encountered
        End If
        
        ' Only update the label if we have a value
        If labelIndex <= 20 And cell.Value <> "" Then
            Area_Dash.Controls("LBLC" & labelIndex).Caption = cell.Value
            Area_Dash.Controls("LBLC" & labelIndex).Visible = True
            Area_Dash.Controls("Imc" & labelIndex).Visible = True
            Area_Dash.Controls("Imd" & labelIndex).Visible = True
            Area_Dash.Controls("LBLCA" & labelIndex).Visible = True
        
            labelIndex = labelIndex + 1
        End If
    Next cell
    
    For i = 1 To 15
        If Area_Dash.Controls("LB" & i).Caption <> "" Then
            Area_Dash.Controls("Im" & i).Visible = True
            Area_Dash.Controls("LB" & i).ForeColor = vbWhite
        End If
        Next i
        
        For i = 1 To 3
        If Area_Dash.Controls("LBL" & i).Caption <> "" Then
            Area_Dash.Controls("Img" & i).Visible = True
            Area_Dash.Controls("LBL" & i).ForeColor = vbWhite
        End If
        Next i
    
      For i = 1 To 20
        Set lblc = Area_Dash.Controls("Imc" & i)
        
        ' ???? ??? ??? ??? label ???? ??? ????
        If lblc.Visible = True Then
            visibleLabelsc = visibleLabelsc + 1
        End If
    Next i
        
           If visibleLabelsc > 8 Then
    Area_Dash.ScrollBar2.Visible = True
    Area_Dash.ScrollBar2.Value = 0
    End If
    
    Area_Dash.N_LB1.Caption = visibleLabelsc
        
End Sub

Private Sub LB_Left_Click()

Dim i As Integer
Dim PG As MultiPage
Dim ctrl As Control
Dim multiPageCount As Integer
Dim pagecount As Integer
    
    'count multi pages on the userform
    multiPageCount = 0
    
    
    For Each ctrl In Area_Dash.Controls
        
        If TypeName(ctrl) = "MultiPage" Then
            multiPageCount = multiPageCount + 1
        End If
    Next ctrl

    For i = 1 To multiPageCount

    If Area_Dash.Controls("MultiPage" & i).Visible = True Then
        pagecount = Area_Dash.Controls("MultiPage" & i).Pages.Count
        
    With Area_Dash.Controls("MultiPage" & i)
        If .Value > 0 Then
            .Value = .Value - 1
        End If
    End With

    Area_Dash.Label36.Caption = Area_Dash.Controls("MultiPage" & i).Value + 1 & " of " & pagecount
    End If

    Next i
End Sub

Private Sub LB_Write_Click()
Dim i As Integer
Dim PG As MultiPage
Dim ctrl As Control
Dim multiPageCount As Integer
Dim pagecount As Integer
    
    'count multi pages on the userform
    multiPageCount = 0
    
    
    For Each ctrl In Area_Dash.Controls
        
        If TypeName(ctrl) = "MultiPage" Then
            multiPageCount = multiPageCount + 1
        End If
    Next ctrl
    
    For i = 1 To multiPageCount
    
    If Area_Dash.Controls("MultiPage" & i).Visible = True Then
        pagecount = Area_Dash.Controls("MultiPage" & i).Pages.Count
    
    With Area_Dash.Controls("MultiPage" & i)
        If .Value < .Pages.Count - 1 Then
            .Value = .Value + 1
        End If
    End With
    Area_Dash.Label36.Caption = Area_Dash.Controls("MultiPage" & i).Value + 1 & " of " & pagecount
    End If
    
    Next i


End Sub

Private Sub LB1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBL1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBL2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBL3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub

Private Sub LBLCA1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub

Private Sub LBLCA11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LBLCA20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB_Write_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub
Private Sub LB_Left_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim hHandCursor As LongPtr
    hHandCursor = LoadCursor(0, IDC_HAND) '  Õ„Ì· √ÌﬁÊ‰… «·Ìœ «·„‘Ì—…
    SetCursor hHandCursor '  ⁄ÌÌ‰ «·„ƒ‘—
End Sub


Private Sub Label25_Click()
Dim lbl As Object
    Dim visibleLabels As Integer
    Dim i As Integer

    visibleLabels = 0 ' ???? ??? labels ???????

    ' ?? ??? labels ?? LB1 ??? LB15 ???? ???? ????? ???? ????
    For i = 1 To 15
        Set lbl = Area_Dash.Controls("LB" & i)
        
        ' ???? ??? ??? ??? label ???? ??? ????
        If lbl.Visible = True And lbl.Caption <> "" Then
            visibleLabels = visibleLabels + 1
        End If
    Next i

    ' ????? ????? ?? ????? visibleLabels
    MsgBox (visibleLabels)
End Sub


Private Sub Label26_Click()
Contracts_Dash.Show
End Sub

Private Sub LB1_Click()

FilterByLabelClick LB1.Caption
hide_buttons
Area_Dash.Im1.Visible = True
Area_Dash.LB1.ForeColor = vbWhite

End Sub

Private Sub LB2_Click()
FilterByLabelClick LB2.Caption
hide_buttons
Area_Dash.Im2.Visible = True
Area_Dash.LB2.ForeColor = vbWhite
End Sub
Private Sub LB3_Click()
FilterByLabelClick LB3.Caption
hide_buttons
Area_Dash.Im3.Visible = True
Area_Dash.LB3.ForeColor = vbWhite
End Sub
Private Sub LB4_Click()
FilterByLabelClick LB4.Caption
hide_buttons
Area_Dash.Im4.Visible = True
Area_Dash.LB4.ForeColor = vbWhite
End Sub
Private Sub LB5_Click()
FilterByLabelClick LB5.Caption
hide_buttons
Area_Dash.Im5.Visible = True
Area_Dash.LB5.ForeColor = vbWhite
End Sub
Private Sub LB6_Click()
FilterByLabelClick LB6.Caption
hide_buttons
Area_Dash.Im6.Visible = True
Area_Dash.LB6.ForeColor = vbWhite
End Sub
Private Sub LB7_Click()
FilterByLabelClick LB7.Caption
hide_buttons
Area_Dash.Im7.Visible = True
Area_Dash.LB7.ForeColor = vbWhite
End Sub
Private Sub LB8_Click()
FilterByLabelClick LB8.Caption
hide_buttons
Area_Dash.Im8.Visible = True
Area_Dash.LB8.ForeColor = vbWhite
End Sub
Private Sub LB9_Click()
FilterByLabelClick LB9.Caption
hide_buttons
Area_Dash.Im9.Visible = True
Area_Dash.LB9.ForeColor = vbWhite
End Sub
Private Sub LB10_Click()
FilterByLabelClick LB10.Caption
hide_buttons
Area_Dash.Im10.Visible = True
Area_Dash.LB10.ForeColor = vbWhite
End Sub
Private Sub LB11_Click()
FilterByLabelClick LB11.Caption
hide_buttons
Area_Dash.Im11.Visible = True
Area_Dash.LB11.ForeColor = vbWhite
End Sub
Private Sub LB12_Click()
FilterByLabelClick LB12.Caption
hide_buttons
Area_Dash.Im12.Visible = True
Area_Dash.LB12.ForeColor = vbWhite
End Sub
Private Sub LB13_Click()
FilterByLabelClick LB13.Caption
hide_buttons
Area_Dash.Im13.Visible = True
Area_Dash.LB13.ForeColor = vbWhite
End Sub
Private Sub LB14_Click()
FilterByLabelClick LB14.Caption
hide_buttons
Area_Dash.Im14.Visible = True
Area_Dash.LB14.ForeColor = vbWhite
End Sub
Private Sub LB15_Click()
FilterByLabelClick LB15.Caption
hide_buttons
Area_Dash.Im15.Visible = True
Area_Dash.LB15.ForeColor = vbWhite
End Sub
Private Sub LBL1_Click()
FilterByLabelClick LBL1.Caption
hide_buttons
Area_Dash.Img1.Visible = True
Area_Dash.LBL1.ForeColor = vbWhite
End Sub
Private Sub LBL2_Click()
FilterByLabelClick LBL2.Caption
hide_buttons
Area_Dash.Img2.Visible = True
Area_Dash.LBL2.ForeColor = vbWhite
End Sub
Private Sub LBL3_Click()
FilterByLabelClick LBL3.Caption
hide_buttons
Area_Dash.Img3.Visible = True
Area_Dash.LBL3.ForeColor = vbWhite
End Sub

Private Sub ScrollBar2_Change()
 Dim i As Integer
    Dim offset As Single

    '  ÕœÌœ ﬁÌ„… «·«‰“·«ﬁ (√Ê «· „—Ì—) „‰ «·‹ ScrollBar
    offset = ScrollBar2.Value

    '  Õ—Ìﬂ «·’Ê— (Im1 ≈·Ï Im15)
    For i = 1 To 20
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ· ’Ê—…
        Me.Controls("Imc" & i).Top = (i - 1) * 28 - offset
    Next i

    '  Õ—Ìﬂ «· ”„Ì«  (LB1 ≈·Ï LB15)
    For i = 1 To 20
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ·  ”„Ì…
        Me.Controls("LBLC" & i).Top = (i - 1) * 28 - offset + 7
    Next i
    
    For i = 1 To 20
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ·  ”„Ì…
        Me.Controls("Imd" & i).Top = (i - 1) * 28 - offset + 3.5
    Next i
        For i = 1 To 20
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ·  ”„Ì…
        Me.Controls("LBLCA" & i).Top = (i - 1) * 28 - offset + 3.5
    Next i
End Sub

Private Sub UserForm_Initialize()
    'Hide all Images and labels from Contracts Section
    Dim lbl As Object
    Dim img As Object
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim zoomlevel As Long
    Dim userwidth As Long
    Dim userheight As Long
    
    screenWidth = Application.UsableWidth
    screenHeight = Application.UsableHeight
    userwidth = Me.Width
    userheight = Me.Height
    zoomlevel = (screenWidth / 1440) * 100

        If zoomlevel < 10 Then
        zoomlevel = 10
        ElseIf zoomlevel > 400 Then
        zoomlevel = 400
        End If
        
        Me.Zoom = zoomlevel
        Me.Width = userwidth * zoomlevel / 100
        Me.Height = userheight * zoomlevel / 100
        
     For i = 1 To 20
        Set lbl = Area_Dash.Controls("LBLC" & i)
        lbl.Visible = False ' Hide the label
        
        Set img = Area_Dash.Controls("Imc" & i)
        img.Visible = False ' Hide the corresponding image
         Set img = Area_Dash.Controls("Imd" & i)
        img.Visible = False ' Hide the corresponding image
    Next i
    '  ⁄ÌÌ‰ «·Õœ «·√ﬁ’Ï ··‹ ScrollBar
    ScrollBar2.Max = 510 - 108
    
    '  ⁄ÌÌ‰ «·Õœ «·√œ‰Ï ··‹ ScrollBar
    ScrollBar1.Min = 0
    ScrollBar2.Min = 0
    
    '  ⁄ÌÌ‰ ÕÃ„ «· „—Ì— ··‹ ScrollBar
    ScrollBar1.SmallChange = 10
    ScrollBar1.Value = 0
    ScrollBar1.LargeChange = 33.95 ' √Ê √Ì ﬁÌ„…  ›÷·Â« ·· Õﬂ„ »ÕÃ„ «· „—Ì—
    
    ScrollBar2.SmallChange = 10
    ScrollBar2.Value = 0
    ScrollBar2.LargeChange = 16
End Sub

Private Sub ScrollBar1_Change()

    Dim i As Integer
    Dim offset As Single


    '  ÕœÌœ ﬁÌ„… «·«‰“·«ﬁ (√Ê «· „—Ì—) „‰ «·‹ ScrollBar
    offset = ScrollBar1.Value

    '  Õ—Ìﬂ «·’Ê— (Im1 ≈·Ï Im15)
    For i = 1 To 15
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ· ’Ê—…
        Me.Controls("Im" & i).Top = (i - 1) * 33.95 - offset
    Next i

    '  Õ—Ìﬂ «· ”„Ì«  (LB1 ≈·Ï LB15)
    For i = 1 To 15
        ' ÷»ÿ Œ«’Ì… «·‹ Top ·ﬂ·  ”„Ì…
        Me.Controls("LB" & i).Top = (i - 1) * 34 - offset + 12.35
    Next i

End Sub
Private Sub Label_Click(lblClicked As MSForms.label)
    Dim lblPair As MSForms.label
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim searchValue As String
    Dim uf As Object

    
    '  ÕœÌœ Ê—ﬁ… «·⁄„·
    Set ws = ThisWorkbook.Sheets("Contracts")
    
    ' «·»ÕÀ ⁄‰ «·‹ label «·„Ê«“Ì ›Ì „Ã„Ê⁄… lblc
    Set lblPair = Me.Controls("lblc" & Mid(lblClicked.Name, 6)) ' «” Œ—«Ã «·—ﬁ„ Ê«” »œ«·Â
    
    ' ﬁ—«¡… «·‰’ «·„ÊÃÊœ ›Ì label «·„Ê«“Ì
    searchValue = lblPair.Caption
    
    ' «·»ÕÀ ⁄‰ «·ﬁÌ„… ›Ì «·ÃœÊ· Contracts_Info ›Ì «·⁄„Êœ D
    Set foundCell = ws.ListObjects("Contracts_Info").ListColumns(4).DataBodyRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' ≈–«  „ «·⁄ÀÊ— ⁄·Ï «·ﬁÌ„…° «› Õ «·‹ UserForm «·√Œ—Ï Ê÷⁄ «·ﬁÌ„…
        Set uf = Contracts_Dash ' «”„ «·›Ê—„
        uf.Controls("Contract_name").Caption = foundCell.Value
        uf.Controls("Project_Box").Value = foundCell.offset(0, -3).Value
        uf.Controls("Speciality_Box").Value = foundCell.offset(0, 10).Value
        uf.Controls("Company_Box").Value = foundCell.offset(0, -1).Value
        uf.Controls("Contract_Date_Box").Value = foundCell.offset(0, 3).Value
        uf.Controls("Contract_End_Box").Value = foundCell.offset(0, 5).Value
        uf.Controls("TextBox3").Value = Format(foundCell.offset(0, 6).Value / 1000000, "#,###.00"" M""")
        Worksheets("Projects_Tab_Pivots").Range("T4").Value = foundCell.offset(0, 7).Value
        Worksheets("Projects_Tab_Pivots").Range("u4").Value = foundCell.offset(0, 8).Value - foundCell.offset(0, 7).Value
        uf.Controls("TextBox1").Value = Format(foundCell.offset(0, 7).Value / 1000000, "#,###.00"" M""")
        uf.Controls("TextBox2").Value = Format(foundCell.offset(0, 8).Value / 1000000, "#,###.00"" M""")
        uf.Controls("Progress_Box").Value = Format(foundCell.offset(0, 9).Value, "0.0%")
        
    Dim chartObj As ChartObject
    Dim tempImagePath As String
    
    Application.Calculate
    ThisWorkbook.Sheets("Projects_Tab_Pivots").Calculate
    
    ' Reference the chart
    Set chartObj = ThisWorkbook.Sheets("Projects_Tab_Pivots").ChartObjects("Chart 1") ' Adjust sheet and chart name
    
    chartObj.Chart.Refresh

    ' Save the chart as a temporary image
    tempImagePath = Environ("TEMP") & "\temp_chart.jpeg"
    chartObj.Chart.Export Filename:=tempImagePath, FilterName:="JPEG"

    ' Load the image into the UserForm's Image control
    Worksheets("Projects_Tab_Pivots").Activate
    Worksheets("Projects_Tab_Pivots").Range("u23").Select
    Worksheets("Dash").Activate
    Contracts_Dash.IMChart.Picture = LoadPicture(tempImagePath)

    ' Optional: Delete the temporary image file
    ' Kill tempImagePath
        
        ' ⁄—÷ «·‹ UserForm
        uf.Show
    Else
        MsgBox "Value not found!", vbExclamation
    End If
End Sub

' √÷› Â–« «·ﬂÊœ ·ﬂ· Label „‰ lblca1 ≈·Ï lblca20
Private Sub lblca1_Click()
    Call Label_Click(Me.LBLCA1)
End Sub

Private Sub lblca2_Click()
    Call Label_Click(Me.LBLCA2)
End Sub

Private Sub lblca3_Click()
    Call Label_Click(Me.LBLCA3)
End Sub
Private Sub lblca4_Click()
    Call Label_Click(Me.LBLCA4)
End Sub
Private Sub lblca5_Click()
    Call Label_Click(Me.LBLCA5)
End Sub
Private Sub lblca6_Click()
    Call Label_Click(Me.LBLCA6)
End Sub
Private Sub lblca7_Click()
    Call Label_Click(Me.LBLCA7)
End Sub
Private Sub lblca8_Click()
    Call Label_Click(Me.LBLCA8)
End Sub
Private Sub lblca9_Click()
    Call Label_Click(Me.LBLCA9)
End Sub
Private Sub lblca10_Click()
    Call Label_Click(Me.LBLCA10)
End Sub
Private Sub lblca11_Click()
    Call Label_Click(Me.LBLCA11)
End Sub
Private Sub lblca12_Click()
    Call Label_Click(Me.LBLCA12)
End Sub
Private Sub lblca13_Click()
    Call Label_Click(Me.LBLCA13)
End Sub
Private Sub lblca14_Click()
    Call Label_Click(Me.LBLCA14)
End Sub
Private Sub lblca15_Click()
    Call Label_Click(Me.LBLCA15)
End Sub
Private Sub lblca16_Click()
    Call Label_Click(Me.LBLCA16)
End Sub
Private Sub lblca17_Click()
    Call Label_Click(Me.LBLCA17)
End Sub
Private Sub lblca18_Click()
    Call Label_Click(Me.LBLCA18)
End Sub
Private Sub lblca19_Click()
    Call Label_Click(Me.LBLCA19)
End Sub
Private Sub lblca20_Click()
    Call Label_Click(Me.LBLCA20)
End Sub





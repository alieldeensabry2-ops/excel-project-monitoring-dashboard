Attribute VB_Name = "Module1"
Sub Lisy2_Show()

If Worksheets("Dash").OLEObjects("ListBox2").Visible = True And Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 98").Visible = True Then
    Worksheets("Dash").OLEObjects("ListBox2").Visible = False
    Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 98").Visible = False
    
Else
    
    Worksheets("Dash").OLEObjects("ListBox2").Visible = True
    Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 98").Visible = True
    
End If

End Sub
Sub Not_Finished()
MsgBox (Worksheets("Projects_Tab_Pivots").Range("B3").Value)

End Sub

Sub Industrial_Area_Dash()
IND_Area.Show
End Sub
Sub Greenhouses_A3()
    
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage1", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A3"
    

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
End Sub
Sub Wholesale_market_A2()
    Dim i As Integer
    Dim multiPageNames As Variant

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Array of the MultiPages to hide
    multiPageNames = Array("MultiPage2", "MultiPage3", "MultiPage4", "MultiPage5", "MultiPage6", "MultiPage7", "MultiPage8", "MultiPage9")

    ' Loop through the array and hide each MultiPage
    For i = LBound(multiPageNames) To UBound(multiPageNames)
        Area_Dash.Controls(multiPageNames(i)).Visible = False
    Next i
    
    ' Call the FilterPivotByAreaAndProjectAndUpdateLabels procedure
    FilterPivotByAreaAndProjectAndUpdateLabels "A2"

    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Sub hide_buttons()
Dim i As Integer
Dim lbl As Object
Dim Image As Object
Dim lbla As Object
Dim Imagea As Object

For i = 1 To 15
Set lbl = Area_Dash.Controls("Im" & i)
Set Image = Area_Dash.Controls("LB" & i)
    Image.ForeColor = vbBlack
    lbl.Visible = False
Next i

For i = 1 To 3
Set lbla = Area_Dash.Controls("Img" & i)
Set Imagea = Area_Dash.Controls("LBL" & i)
    Imagea.ForeColor = vbBlack
    lbla.Visible = False
Next

End Sub

Sub FilterPivotByAreaAndProjectAndUpdateLabels(filterCell As String)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim projectName As String
    Dim i As Integer
    Dim lbl As Object
    Dim img As Object
    Dim resultRange As Range
    Dim resultValues As Variant
    Dim labelIndex As Integer
    Dim visibleLabels As Integer
    Dim chartObj As ChartObject
    Dim tempImagePath As String
    Dim Y As Integer

    ' Disable screen updating and automatic calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Set the worksheet that contains the Pivot Tables
    Set ws = ThisWorkbook.Sheets("Projects_Tab_pivots")
    
    ' Get the value from cell filterCell to use for filtering the Project field in the second Pivot Table
    projectName = ws.Range(filterCell).Value
    ws.Range("B2").Value = projectName
    
    ' Filter the PivotTable2 (second Pivot Table)
    Set pt = ws.PivotTables("PivotTable2")
    pt.PivotFields("Project").ClearAllFilters
    pt.PivotFields("Project").CurrentPage = projectName

    ' Get resultRange from PivotTable2, and store the values in an array for faster processing
    Set resultRange = ws.Range("K6:K" & ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
    resultValues = resultRange.Value ' Store values in array

    ' Ensure resultValues is a 2D array
    If IsArray(resultValues) Then
        ' Update the label on the UserForm with the value from C6
        Area_Dash.Project_name.Caption = projectName
        
        ' Loop through the result values and update the labels on the UserForm
        labelIndex = 1
        For i = 1 To UBound(resultValues, 1)
            If InStr(resultValues(i, 1), "Grand Total") > 0 Then Exit For
            If labelIndex <= 15 Then
                Area_Dash.Controls("LB" & labelIndex).Caption = resultValues(i, 1)
                labelIndex = labelIndex + 1
            Else
                Exit For
            End If
        Next i
    End If

    ' Hide empty labels and their corresponding images (LB1 to LB15)
    For i = 1 To 15
        Set lbl = Area_Dash.Controls("LB" & i)
        Set img = Area_Dash.Controls("Im" & i)
        If lbl.Caption = "" Then
            lbl.Visible = False
            img.Visible = False
        Else
            lbl.Visible = True
            img.Visible = True
        End If
    Next i

    ' Now filter the first PivotTable (PivotTable1) based on the value in C6 (for the "Project" field)
    Set pt = ws.PivotTables("PivotTable1")
    pt.PivotFields("Project").ClearAllFilters
    pt.PivotFields("Project").CurrentPage = projectName

    ' Get resultRange from PivotTable1, and store the values in an array for faster processing
    Set resultRange = ws.Range("G6:G" & ws.Cells(ws.Rows.Count, "G").End(xlUp).Row)
    resultValues = resultRange.Value ' Store values in array

    ' Ensure resultValues is a 2D array
    If IsArray(resultValues) Then
        ' Update labels (LBL1 to LBL3)
        labelIndex = 1
        For i = 1 To UBound(resultValues, 1)
            If InStr(resultValues(i, 1), "Grand Total") > 0 Then Exit For
            If labelIndex <= 3 Then
                Area_Dash.Controls("LBL" & labelIndex).Caption = resultValues(i, 1)
                labelIndex = labelIndex + 1
            Else
                Exit For
            End If
        Next i
    End If

    ' Hide empty labels and their corresponding images (LBL1 to LBL3)
    For i = 1 To 3
        Set lbl = Area_Dash.Controls("LBL" & i)
        Set img = Area_Dash.Controls("Img" & i)
        If lbl.Caption = "" Then
            lbl.Visible = False
            img.Visible = False
        Else
            lbl.Visible = True
            img.Visible = True
        End If
    Next i

    ' Update scroll bar visibility based on the number of visible labels
    visibleLabels = 0
    For i = 1 To 15
        Set lbl = Area_Dash.Controls("LB" & i)
        If lbl.Visible = True And lbl.Caption <> "" Then
            visibleLabels = visibleLabels + 1
        End If
    Next i
    Area_Dash.ScrollBar1.Max = (visibleLabels * 36) - 108
    If visibleLabels <= 3 Then Area_Dash.ScrollBar1.Visible = False
    
    'set the project analysis values from the pivot
    
    ws.PivotTables("Project_Details").PivotFields("Project").ClearAllFilters
    ws.PivotTables("Project_Details").PivotFields("Project").CurrentPage = projectName
    
    Area_Dash.D_Box1.Caption = Format(Worksheets("Projects_Tab_Pivots").Range("AB6").Value / 1000000, "#,###"" M""")
    Area_Dash.D_Box2.Caption = Format(Worksheets("Projects_Tab_Pivots").Range("Ac6").Value / 1000000, "#,###"" M""")
    Area_Dash.D_Box3.Caption = Format(Worksheets("Projects_Tab_Pivots").Range("AD6").Value / 1000000, "#,###"" M""")
    Area_Dash.D_Box4.Caption = Format(Worksheets("Projects_Tab_Pivots").Range("AE6").Value, "0.0%")
    
    
    
    
    ' Process the chart (save as image and update the UserForm)
    Application.Calculate
    ThisWorkbook.Sheets("Projects_Tab_Pivots").Calculate
    
    Set chartObj = ThisWorkbook.Sheets("Projects_Tab_Pivots").ChartObjects("Chart 6")
    
    chartObj.Chart.Refresh
    
    tempImagePath = Environ("TEMP") & "\temp_chart.jpeg"
    chartObj.Chart.Export Filename:=tempImagePath, FilterName:="JPEG"
    Area_Dash.P_Bar.Picture = LoadPicture(tempImagePath)

    ' Update the multi-page control visibility
    Dim multiPageCount As Integer
    Dim pagecount As Integer
    
    'count multi pages on the userform
    multiPageCount = 0
    
    
    For Each ctrl In Area_Dash.Controls
        
        If TypeName(ctrl) = "MultiPage" Then
            multiPageCount = multiPageCount + 1
        End If
    Next ctrl

    For Y = 1 To multiPageCount

    If Area_Dash.Controls("MultiPage" & Y).Visible = True Then
            Area_Dash.Controls("MultiPage" & Y).Value = 0
            Area_Dash.Label36.Caption = "1 of " & Area_Dash.Controls("MultiPage" & Y).Pages.Count
        End If
    Next Y
    
    'page count
    
    
    'show all the contracts in the projects tab
    show_inDash
    
    'scrollbar2 visibility
     visibleLabelsc = 0

    
    For i = 1 To 20
        Set lblc = Area_Dash.Controls("Imc" & i)
        
        
        If lblc.Visible = True Then
            visibleLabelsc = visibleLabelsc + 1
        End If
    Next i

    If visibleLabelsc <= 8 Then
    Area_Dash.ScrollBar2.Visible = False
    End If
        Area_Dash.N_LB1.Caption = visibleLabelsc
        
     
    ' Show the UserForm
    Area_Dash.Show

    ' Re-enable screen updating and calculation after all operations are done
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub



Sub hidemultipages()

    Dim ctrl As Control
    Dim multiPageCount As Integer
    Dim i As Integer
    multiPageCount = 0
    
    
    For Each ctrl In Area_Dash.Controls
        
        If TypeName(ctrl) = "MultiPage" Then
            multiPageCount = multiPageCount + 1
        End If
    Next ctrl
    

    
    For i = 1 To multiPageCount
        Area_Dash.Controls("MultiPage" & i).Visible = False
    Next i
    

    

End Sub


' New Sub for handling label clicks:
Sub FilterByLabelClick(selectedLabel As String)
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
    Dim visibleLabelsc As Integer
    Dim X As Integer
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
    pt.PivotFields("Company Name").CurrentPage = selectedLabel
    
    ' Get the results from the PivotTable in column O
    Set resultRange = ws.Range("O6:O" & ws.Cells(ws.Rows.Count, "O").End(xlUp).Row)
    X = resultRange.Count
    
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
     visibleLabelsc = 0 ' ???? ??? labels ???????

    ' ?? ??? labels ?? LB1 ??? LB15 ???? ???? ????? ???? ????
    For i = 1 To 20
        Set lblc = Area_Dash.Controls("Imc" & i)
        
        ' ???? ??? ??? ??? label ???? ??? ????
        If lblc.Visible = True Then
            visibleLabelsc = visibleLabelsc + 1
        End If
    Next i

    If visibleLabelsc <= 8 Then
    Area_Dash.ScrollBar2.Visible = False
    Area_Dash.ScrollBar2.Value = 0
    End If
    Area_Dash.N_LB1.Caption = X - 1
End Sub


Sub show_inDash()
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
    

End Sub
Sub PlayVideo()
    Dim wsCurrent As Worksheet
    Dim wsVideo As Worksheet
    Dim obj As OLEObject
    
    Application.ScreenUpdating = False
        
    ' ??? ?????? ???????
    Set wsCurrent = ActiveSheet
    
    ' ????? ?????? ???? ????? ??? ????????
    Set wsVideo = ThisWorkbook.Sheets("Projects_Tab_pivots")
    
    ' ????? ?????? ????? ??? ???? ?? Selection Panel
    Set obj = wsVideo.OLEObjects("Object 5")
    
    ' ????? ??????
    obj.Verb xlPrimary
    
    ' ?????? ??? ?????? ???????
    wsCurrent.Activate
End Sub





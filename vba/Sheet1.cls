VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Private Sub Search_Box_Change()
    Dim ws As Worksheet
    Dim searchSheet As Worksheet
    Dim searchTerm As String
    Dim cell As Range
    Dim suggestions As Collection
    Dim suggestion As Variant
    Dim activeXListBox As Object
    Dim itemCount As Integer
    Dim maxHeight As Integer
    Dim itemHeight As Integer
    Dim searchColumn As String
    

    ' Set worksheets
    Set ws = ThisWorkbook.Sheets("dash")
    Set searchSheet = ThisWorkbook.Sheets("Contracts")
    Set activeXListBox = ws.OLEObjects("ListBox1").Object ' Assign ActiveX ListBox
    Set suggestions = New Collection ' Create collection for suggestions
    
    
    ' hide the X icon
    ws.Shapes("Graphic 109").Visible = True
    
    ' Get the text entered in the Search Box
    searchTerm = ws.OLEObjects("Search_Box").Object.Text

    
    Select Case ThisWorkbook.Sheets("Projects_Tab_Pivots").Range("A26").Value
        Case 0
            searchColumn = "D"
        Case 1
            searchColumn = "A"
        Case 2
            searchColumn = "C"
        Case Else
            MsgBox "Invalid value in A26. Please enter 0, 1, or 2.", vbExclamation
            Exit Sub
    End Select

    ' Clear previous items in ListBox
    activeXListBox.Clear

    ' Count the number of suggestions that match the search term
    itemCount = 0

    ' Search for matches
    If searchTerm <> "" Then
        For Each cell In searchSheet.Range(searchColumn & "2:" & searchColumn & searchSheet.Cells(searchSheet.Rows.Count, searchColumn).End(xlUp).Row)
            If InStr(1, cell.Value, searchTerm, vbTextCompare) > 0 Then
                On Error Resume Next
                suggestions.Add cell.Value, CStr(cell.Value) ' Add unique suggestions
                On Error GoTo 0
                itemCount = itemCount + 1 ' Increment count
            End If
        Next cell

        ' Add the suggestions to the ListBox
        For Each suggestion In suggestions
            activeXListBox.AddItem suggestion
        Next suggestion
    End If

    ' Adjust the ListBox height based on the number of items
    itemHeight = 15 ' Height per item in ListBox
    maxHeight = 70 ' Maximum height for the ListBox

    ' Set the ListBox height based on the number of items
    If itemCount > 0 Then
        If itemCount * itemHeight <= maxHeight Then
            ws.OLEObjects("ListBox1").Height = itemCount * itemHeight ' Adjust height based on item count
            ws.Shapes("Rectangle: Top Corners Rounded 99").Height = itemCount * (itemHeight + 6)
        Else
            ws.OLEObjects("ListBox1").Height = maxHeight ' Set maximum height
            ws.Shapes("Rectangle: Top Corners Rounded 99").Height = maxHeight + 10
        End If
    End If

    ' Toggle visibility of ListBox and Rectangle based on suggestions count
    If suggestions.Count < 1 Then
        ws.OLEObjects("ListBox1").Visible = False
        ws.Shapes("Rectangle: Top Corners Rounded 99").Visible = False
    Else
        ws.OLEObjects("ListBox1").Visible = True
        ws.Shapes("Rectangle: Top Corners Rounded 99").Visible = True
    End If

    ' ?????? ?? ????? ??? Top
    If ws.Shapes("Rectangle: Top Corners Rounded 99").Top <> 95 Then
        ws.Shapes("Rectangle: Top Corners Rounded 99").Top = 95
    End If
    
    If Worksheets("Dash").OLEObjects("Search_Box").Object.Text = "" Then
            Worksheets("Dash").Shapes("Graphic 109").Visible = False
    End If
    
End Sub

Private Sub ListBox1_Click()
    Dim ws As Worksheet
    Dim activeXListBox As Object
    Dim selectedItem As String
    
    Set ws = ThisWorkbook.Sheets("dash")

    On Error Resume Next
    Set activeXListBox = ws.OLEObjects("ListBox1").Object
    On Error GoTo 0
    
    If activeXListBox Is Nothing Then
        Exit Sub
    End If
    
    ' Ensure an item is selected
    If activeXListBox.ListIndex <> -1 Then
        ' Get the selected item from ListBox
        selectedItem = activeXListBox.List(activeXListBox.ListIndex)
        Worksheets("Projects_Tab_Pivots").Range("A29").Value = selectedItem
        
        ' Remove any space between a number and a word
        selectedItem = RemoveSpaceBetweenNumberAndWord(selectedItem)
        
        ' Set the cleaned value as the value of the Search Box
        ws.OLEObjects("Search_Box").Object.Text = selectedItem
    End If

    ' Hide the ListBox and other shapes
    ThisWorkbook.Sheets("dash").OLEObjects("ListBox1").Visible = False
    Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 99").Visible = False
End Sub

' Function to remove space between numbers and words
Function RemoveSpaceBetweenNumberAndWord(inputText As String) As String
    Dim i As Integer
    Dim result As String
    result = inputText

    ' Loop through each character in the string
    For i = 1 To Len(result) - 1
        ' Check if the current character is a number and the next character is a space
        If IsNumeric(Mid(result, i, 1)) And Mid(result, i + 1, 1) = " " Then
            ' Remove the space between the number and the next character
            result = Left(result, i) & Mid(result, i + 2)
        End If
    Next i

    ' Return the cleaned string
    RemoveSpaceBetweenNumberAndWord = result
End Function

Private Sub ListBox2_Click()
    Dim ws As Worksheet
    Dim ShapeX As Object
    
    Set ws = ThisWorkbook.Sheets("dash")
    Set ShapeX = ws.OLEObjects("ListBox2").Object

    ' Ensure an item is selected
    If ShapeX.ListIndex <> -1 Then
        ' Get the selected item and set it as the value of the Search Box
        ws.Shapes("FL_Cap").OLEFormat.Object.Text = ShapeX.List(ShapeX.ListIndex)
        Worksheets("Projects_Tab_Pivots").Range("A26").Value = ShapeX.ListIndex
    End If
    ThisWorkbook.Sheets("dash").OLEObjects("ListBox2").Visible = False
    Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 98").Visible = False
    
End Sub

Sub Search_Selection()
    
    If Worksheets("Projects_Tab_Pivots").Range("A26").Value = 0 Then
        run_Contract
        
    ElseIf Worksheets("Projects_Tab_Pivots").Range("A26").Value = 1 Then
        If Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A2").Value Then
            Wholesale_market_A2
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A3").Value Then
                Greenhouses_A3
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A4").Value Then
                    Call IND_Area.Label1_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A5").Value Then
                    Call IND_Area.Label2_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A6").Value Then
                    Call IND_Area.Label4_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A7").Value Then
                    Call IND_Area.Label3_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A8").Value Then
                    Call IND_Area.Label5_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A9").Value Then
                    Call IND_Area.Label6_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A10").Value Then
                    Call IND_Area.Label7_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A11").Value Then
                    Call IND_Area.Label8_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A12").Value Then
                    Call IND_Area.Label9_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A13").Value Then
                    Call IND_Area.Label10_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A14").Value Then
                    Call IND_Area.Label11_Click
            ElseIf Worksheets("Projects_Tab_Pivots").Range("A29").Value = Worksheets("Projects_Tab_Pivots").Range("A15").Value Then
                    Call IND_Area.Label12_Click
        End If
    End If
    
        'delete the value of the searchbox
            Worksheets("Dash").OLEObjects("Search_Box").Object.Text = ""
End Sub
Sub run_Contract()
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim searchValue As String
    Dim uf As Object
    
    '  ÕœÌœ Ê—ﬁ… «·⁄„·
    Set ws = ThisWorkbook.Sheets("Contracts")
    
    
    ' ﬁ—«¡… «·‰’ «·„ÊÃÊœ ›Ì label «·„Ê«“Ì
    searchValue = Worksheets("Projects_Tab_Pivots").Range("A29").Value
    
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
    
    ' Reference the chart
    Set chartObj = ThisWorkbook.Sheets("Projects_Tab_Pivots").ChartObjects("Chart 1") ' Adjust sheet and chart name

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

Sub cancel_Search()


    ThisWorkbook.Sheets("dash").OLEObjects("ListBox2").Visible = False
    Worksheets("Dash").Shapes("Rectangle: Top Corners Rounded 98").Visible = False
    Worksheets("Dash").OLEObjects("Search_Box").Object.Text = ""
    Worksheets("Dash").Shapes("Graphic 109").Visible = False
End Sub

Private Sub Search_Box_GotFocus()
    If Worksheets("Dash").OLEObjects("Search_Box").Object.Text <> "" Then
        Worksheets("Dash").Shapes("Graphic 109").Visible = True
    End If
    
    
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Workbook_Open()

    Dim screenWidth As Long, screenHeight As Long
    Dim zoomlevel As Integer
    
    screenWidth = Application.UsableWidth
    screenHeight = Application.UsableHeight
    
    zoomlevel = (screenWidth / 1440) * 100

        If zoomlevel < 10 Then
        zoomlevel = 10
        ElseIf zoomlevel > 400 Then
        zoomlevel = 400
        End If
        
    Worksheets("Dash").ScrollArea = "A1:Ac53"
         
    '  ÿ»Ìﬁ „” ÊÏ «· ﬂ»Ì— ⁄·Ï «·Ê—ﬁ… «·‰‘ÿ…
    ThisWorkbook.Sheets("Dash").Activate
    ActiveWindow.Zoom = zoomlevel
    Application.DisplayFullScreen = True
End Sub

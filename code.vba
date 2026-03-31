Sub SmartContentsManager()

    Dim ws As Worksheet
    Dim contentSheet As Worksheet
    Dim btn As Object
    Dim i As Integer
    Dim sheetExists As Boolean

    ' Check if Contents sheet exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Contents" Then
            sheetExists = True
            Exit For
        End If
    Next ws

    ' Delete only if exists
    If sheetExists = True Then
        Application.DisplayAlerts = False
        Worksheets("Contents").Delete
        Application.DisplayAlerts = True
    End If

    ' Create new Contents sheet at first position
    Set contentSheet = Worksheets.Add(Before:=Worksheets(1))
    contentSheet.Name = "Contents"

    ' Title formatting
    With contentSheet.Range("A1")
        .Value = "Workbook Contents"
        .Font.Bold = True
        .Font.Size = 16
    End With

    i = 3

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name <> "Contents" Then
            
            ' Add hyperlink
            contentSheet.Hyperlinks.Add _
                Anchor:=contentSheet.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name

            i = i + 1

            ' Delete old button if exists
            On Error Resume Next
            ws.Shapes("BackButton").Delete
            On Error GoTo 0

            ' Add Back button
            Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 150, 30)

            With btn
                .Name = "BackButton"
                .TextFrame.Characters.Text = "? Back to Contents"
                .OnAction = "GoToContents"
                .Fill.ForeColor.RGB = RGB(0, 102, 204)
                .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            End With

        End If
        
    Next ws

    MsgBox "? Smart Contents Page Created Successfully!", vbInformation

End Sub


Sub GoToContents()

    Dim ws As Worksheet
    Dim found As Boolean
    
    found = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Contents" Then
            ws.Activate
            found = True
            Exit For
        End If
    Next ws
    
    If found = False Then
        MsgBox "? Contents sheet not found!", vbExclamation
    End If

End Sub


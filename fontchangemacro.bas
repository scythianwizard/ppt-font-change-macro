Private Sub CommandButton1_Click()
' This macro changes the font to the specified font in "sFontName"
' This works for: First level shapes, shapes in groups and smart art, and tables
' This will not work for: Charts
Dim myValue As Variant
    Dim oSl As Slide ' Variable(Dimension) for Slide
    Dim oSh As Shape ' Variable(Dimension) Alias for shape
    Dim oSh2 As Shape ' Variable(Dimension) for another shape, here we are using this to go within shapes that are grouped
    Dim oTbl As Table ' Variable(Dimension) for the table
    Dim lRow As Long ' Variable(Dimension) for the rows of the table
    Dim lCol As Long ' Variable(Dimension) for the columns of the table
    Dim sFontName As String ' ' Variable(Dimension) for the font
    
    If MsgBox("Are you sure you want to change the font?", vbYesNo) = vbNo Then Exit Sub

    sFontName = InputBox("Please enter the font name", "Font name", "Arial")
        'Quick Copy-Paste ref: Arial , Verdana, Meiryo, Times New Roman

    With ActivePresentation ' Run this part in the active file
        For Each oSl In .Slides
        ' loop through all the slides
            For Each oSh In oSl.Shapes
            ' For each Shape among the shapes of the slide
                If oSh.HasTextFrame Then
                ' Check if the shape has text frame
                    If oSh.TextFrame.HasText Then
                    ' Check if the text frame has text
                        oSh.TextFrame.TextRange.Font.Name = sFontName
                        ' Change the font of the text
                    End If ' (if text frame has text)
                ' (First level text frame)
                
                ' What if it is not a first level text frame, but has shapes within it?
                ElseIf oSh.Type = msoGroup Then
                ' Check if the type of the shape is group (if the shapes are grouped)
                    For Each oShp2 In oSh.GroupItems
                    ' For each (second level) shape within the grouped items (within the first level shape)
                        If oShp2.HasTextFrame Then
                        ' If the (second level) shape has text frame in it
                            If oShp2.TextFrame.HasText Then
                                oShp2.TextFrame.TextRange.Font.Name = sFontName
                               ' Change the font of the text
                            End If '(text frame has text)
                        End If '((second level) shape text frame)
                    Next ' ( For loop for group)
                
                
                ' Now, what if there are tables with text?
                ElseIf oSh.HasTable Then
                ' Check if the shape contains a table
                    Set oTbl = oSh.Table
                    ' Set the table in the table dimension
                    For lRow = 1 To oTbl.Rows.Count
                    ' Loop for rows ranging from first to the last rows in the table
                      For lCol = 1 To oTbl.Columns.Count
                      ' Next, loop for all the columns within the current row
                            With oTbl.Cell(lRow, lCol).Shape.TextFrame.TextRange
                                .Font.Name = sFontName
                            End With
                            ' Repace the font with the above font
                      Next ' (Column loop)
                    Next ' (Row loop)
                ' Next, Smart Art
                ElseIf oSh.HasSmartArt Then
                    For Each oShp2 In oSh.GroupItems
                    ' For the shapes within the SmartArt
                        If oShp2.HasTextFrame Then
                            ' If the (second level) shape has text frame in it
                            If oShp2.TextFrame.HasText Then
                                oShp2.TextFrame.TextRange.Font.Name = sFontName
                                ' Change the font of the text
                            End If '(text frame has text)
                        End If '((second level) shape text frame)
                    Next ' Loop for shapes within the chart
                End If
            Next
            ' Next shape
        Next
        ' Next slide
    End With
    MsgBox cnt & "Font changed!"
End Sub

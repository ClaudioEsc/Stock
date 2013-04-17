Attribute VB_Name = "modFlexgrid"
Option Explicit

Public Enum EGridCellProperty
    gcpAlignment
    gcpFontName
    gcpFontSize
    gcpFontBold
    gcpForeColor
    gcpBackColor
End Enum

Public Enum EGridColAlign
    gcaLeft
    gcaRight
    gcaCenter
End Enum

Public Sub FillGrid(ByRef grd As MSFlexGrid, _
                    ByRef rs As ADODB.Recordset, _
                    Optional ByVal PreserveRow As Boolean = False, _
                    Optional ByVal AutoSizeCols As Boolean = False)
                
    Dim Col         As Long
    Dim Row         As Long
    Dim FxRow       As Long
    Dim FxCol       As Long
    Dim PrevRow     As Long
    
On Error GoTo Catch
    With grd
        FxRow = .FixedRows
        FxCol = .FixedCols
        
        If PreserveRow Then
            PrevRow = .Row
        Else
            PrevRow = 1
        End If
        
        .Redraw = False
        .Rows = FxRow

        If Not EmptyRS(rs) Then
            .Rows = FxRow + rs.RecordCount
            
            rs.MoveFirst
            
            For Row = FxRow To .Rows - 1
                For Col = FxCol To .Cols - 1
                    .TextMatrix(Row, Col) = rs.Collect(Col - FxCol) & vbNullString
                Next
                
                rs.MoveNext
            Next
        End If
    End With
    
Finally:
    grd.Redraw = True
    
    Exit Sub
Catch:
    ErrorReport "FillGrid", "modFlexgrid"
    Resume Finally
End Sub

Public Function ExportarExcel(ByVal FileName As String, ByRef fg As Object, ByVal Title As String) As Boolean
    Dim Excel   As Object
    Dim Book    As Object
    Dim Sheet   As Object
    Dim Row     As Long
    Dim Col     As Long

On Error GoTo ErrHandler
    
    Set Excel = CreateObject("Excel.Application")
    Set Book = Excel.Workbooks.Add
    Set Sheet = Book.Worksheets.Add

    With Sheet
        For Row = 0 To fg.Rows - 1
            For Col = fg.FixedCols To fg.Cols - 1
                With .Cells(Row + 1, Col + 1 - fg.FixedCols)
                    .Font.Bold = (Row = 0)
                    .Value = fg.TextMatrix(Row, Col)
                End With
            Next
        Next

        With .PageSetup
            .LeftHeader = Title
            .RightHeader = "&D, &T"
            .CenterFooter = "Página &P de &N"
        End With
    End With
    
    Book.Close True, FileName
    Excel.Quit

    ExportarExcel = True
 
ResumePoint:
    Set Excel = Nothing
    Set Book = Nothing
    Set Sheet = Nothing
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbCritical, Err.Number
    Resume ResumePoint
End Function

Public Sub AutoSize(ByRef grd As MSFlexGrid, _
                    Optional ByVal Col1 As Long = 0, _
                    Optional ByVal Col2 As Long = 0)
    Dim Col As Long
    Dim Row As Long
    Dim TopRow As Long
    Dim ColWidth() As Single
    Dim CellWidth  As Single
    
    With grd
        If .Rows > 0 Then
            ReDim ColWidth(0 To .Cols - 1)
            
            For Col = 0 To .Cols - 1
                ColWidth(Col) = .Parent.TextWidth(.TextMatrix(0, Col))
            Next
            
            'Solamente se ajustan 200 filas.
            If .Rows > 200 Then
                TopRow = 200
            Else
                TopRow = .Rows - 1
            End If
            
            If Col1 = 0 And Col2 = 0 Then
                Col1 = 0
                Col2 = .Cols - 1
            ElseIf Col1 <> 0 And Col2 = 0 Then
                Col2 = Col1
            End If
    
            If Col2 = 0 Then Col2 = .Cols - 1
                        
            For Row = 1 To TopRow
                For Col = Col1 To Col2
                    CellWidth = .Parent.TextWidth(.TextMatrix(Row, Col))
                    If ColWidth(Col) < CellWidth Then
                        ColWidth(Col) = CellWidth
                    End If
                Next
            Next
            
            For Col = Col1 To Col2
                .ColWidth(Col) = ColWidth(Col) + 300
            Next
        End If
    End With
End Sub

Public Sub SelectRow(ByRef grd As MSFlexGrid, Optional ByVal Row As Long = 1)
On Error Resume Next
    With grd
        .Redraw = False
        
        If .Rows > .FixedRows Then
            If Row < .FixedRows Then
                Row = 1
            ElseIf Row > .Rows Then
                Row = .Rows - 1
            End If
            
            If Not .RowIsVisible(Row) Then
                .TopRow = Row
            End If
            
            .Row = Row
            .Col = .FixedCols
            .RowSel = Row
            .ColSel = .Cols - 1
        Else
            .Row = 0
            .Col = 0
        End If
        
        .Redraw = True
    End With
End Sub

Public Sub CellProperty(ByRef grd As MSFlexGrid, _
                        ByVal Property As EGridCellProperty, _
                        ByVal NewValue As Variant, _
                        ByVal FromRow As Long, _
                        ByVal FromCol As Long, _
                        Optional ByVal ToRow As Long = -1, _
                        Optional ByVal ToCol As Long = -1)

    Dim PrevRow As Long
    Dim PrevCol As Long
    Dim PrevFillStyle As Long
    Dim PrevRedraw As Boolean
        
    With grd
        If .Rows = 0 Then Exit Sub
        If .Cols = 0 Then Exit Sub
        If FromRow > .Rows - 1 Then Exit Sub
        If FromCol > .Cols - 1 Then Exit Sub
        
        PrevRow = .Row
        PrevCol = .Col
        PrevFillStyle = .FillStyle
        PrevRedraw = .Redraw
        
        .Redraw = False
        .FillStyle = flexFillRepeat
        
        If FromRow = -1 Then
            FromRow = 0
        Else
            If FromRow >= 0 And FromRow < .Rows - 1 Then
                .Row = FromRow
            End If
        End If
        
        .Row = FromRow
        .Col = FromCol
        
        If ToRow = -1 Then
            .RowSel = FromRow
        Else
            .RowSel = ToRow
        End If
        
        
        If ToCol = -1 Then
            .ColSel = FromCol
        Else
            .ColSel = ToCol
        End If
        
        Select Case Property
            Case gcpAlignment
                .CellAlignment = NewValue
                
            Case gcpFontBold
                .CellFontBold = NewValue
                
            Case gcpFontName
                .CellFontName = NewValue
                
            Case gcpFontSize
                .CellFontSize = NewValue
                
            Case gcpForeColor
                .CellForeColor = NewValue
                
            Case gcpBackColor
                .CellBackColor = NewValue
        End Select
        
        .Row = PrevRow
        .Col = PrevCol
        .FillStyle = PrevFillStyle
        .Redraw = PrevRedraw
    End With
End Sub

Public Sub RowProperty(ByRef grd As MSFlexGrid, _
                       ByVal Property As EGridCellProperty, _
                       ByVal NewValue As Variant, _
                       ByVal Row As Long)
    CellProperty grd, Property, NewValue, Row, grd.FixedCols, Row, grd.Cols - 1
End Sub

Public Sub ColProperty(ByRef grd As MSFlexGrid, _
                       ByVal Property As EGridCellProperty, _
                       ByVal NewValue As Variant, _
                       ByVal Col As Long)
    CellProperty grd, Property, NewValue, grd.FixedRows, Col, grd.Rows - 1, Col
End Sub

Public Sub InitGridCol(ByRef grd As MSFlexGrid, _
                       ByVal Col As Long, _
                       ByVal Text As String, _
                       Optional ByVal Width As Long = 1000, _
                       Optional ByVal Align As EGridColAlign)
    With grd
        .TextMatrix(0, Col) = Text
        .ColWidth(Col) = Width
        
        Select Case Align
            Case gcaCenter: .ColAlignment(Col) = flexAlignCenterCenter
            Case gcaRight:  .ColAlignment(Col) = flexAlignRightCenter
            Case gcaLeft:   .ColAlignment(Col) = flexAlignLeftCenter
        End Select
    End With
End Sub
                       


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

Public Function GridExportExcel(ByRef grd As MSFlexGrid, _
                                ByVal FileName As String, _
                                ByVal Title As String) As Boolean
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
        For Row = 0 To grd.Rows - 1
            For Col = grd.FixedCols To grd.Cols - 1
                With .Cells(Row + 1, Col + 1 - grd.FixedCols)
                    .Font.Bold = (Row = 0)
                    .Value = grd.TextMatrix(Row, Col)
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

    GridExportExcel = True
 
ResumePoint:
    Set Excel = Nothing
    Set Book = Nothing
    Set Sheet = Nothing
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbCritical, Err.Number
    Resume ResumePoint
End Function

Public Sub GridAutoSize(ByRef grd As MSFlexGrid, _
                        Optional ByVal Column As Long = -1)
    Dim Col     As Long
    Dim Row     As Long
    Dim ToRow   As Long
    Dim FromCol As Long
    Dim ToCol   As Long
    Dim ColWidth() As Single
    Dim CellWidth  As Single
    
    With grd
        If .Rows > 0 And .Cols > 0 Then
            ReDim ColWidth(0 To .Cols - 1)
            
            For Col = 0 To .Cols - 1
                ColWidth(Col) = .Parent.TextWidth(.TextMatrix(0, Col))
            Next
            
            'Solamente se ajustan 200 filas.
            If .Rows > 200 Then
                ToRow = 200
            Else
                ToRow = .Rows - 1
            End If
            
            If Column = -1 Then
                FromCol = 0
                ToCol = .Cols - 1
            Else
                FromCol = Column
                ToCol = Column
            End If
            
            For Row = 1 To ToRow
                For Col = FromCol To ToCol
                    CellWidth = .Parent.TextWidth(.TextMatrix(Row, Col))
                    
                    If ColWidth(Col) < CellWidth Then
                        ColWidth(Col) = CellWidth
                    End If
                Next
            Next
            
            For Col = FromCol To ToCol
                .ColWidth(Col) = ColWidth(Col) + 300
            Next
        End If
    End With
End Sub

Public Sub GridSelectRow(ByRef grd As MSFlexGrid, Optional ByVal Row As Long = 1)
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

Public Sub GridInitCol(ByRef grd As MSFlexGrid, _
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
                       
Public Function GridFindRow(ByRef grd As MSFlexGrid, _
                            ByVal Item As String, _
                            Optional ByVal Row As Long = -1, _
                            Optional ByVal Col As Long = -1, _
                            Optional ByVal FullMatch As Boolean = True) As Long
    Dim i As Long
    Dim j As Long
    Dim RetRow As Long
    
    RetRow = -1
    
    With grd
        If Row = -1 Then
            Row = .FixedRows
        End If
        
        ' Busqueda por RowData
        If Col = -1 Then
            For i = Row To .Rows - 1
                If .RowData(i) = Val(Item) Then
                    RetRow = i
                    Exit For
                End If
            Next
        Else
            ' Busqueda de palabra completa
            If FullMatch Then
                For i = Row To .Rows - 1
                    If .TextMatrix(i, Col) = Item Then
                        RetRow = i
                        Exit For
                    End If
                Next
            Else
            ' Busqueda por porcion de palabra
                For i = Row To .Rows - 1
                    If InStr(1, .TextMatrix(i, Col), Item) Then
                        RetRow = i
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    GridFindRow = RetRow
End Function

Public Sub GridRemoveRow(ByRef grd As MSFlexGrid, ByVal Row As Long)
    With grd
        If .Rows > .FixedRows Then
            If .Rows - .FixedRows = 1 Then
                .Rows = 1
            Else
                .RemoveItem Row
            End If
        End If
    End With
End Sub

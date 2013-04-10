Attribute VB_Name = "modFlexgrid"
Option Explicit

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
'    grd.Row = PrevRow
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

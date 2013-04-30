Attribute VB_Name = "modADO"
Option Explicit

Public Enum EFieldType
    fdtString
    fdtLong
    fdtDecimal
    fdtCurrency
    fdtDate
    fdtBoolean
    fdtOther
End Enum

Private m_Cn        As ADODB.Connection
Private m_InTrans   As Boolean

Public Function InitConnection() As Boolean
    Dim cnString As String
    Dim PathDB  As String

    cnString = "Driver=SQLite3 ODBC Driver;Database=" & ReadINI("db", "path", App.Path & "\base.s3db")
    
On Error GoTo Catch
    Set m_Cn = New ADODB.Connection
    
    With m_Cn
        .CursorLocation = adUseClient
        .Open cnString
    End With

    InitConnection = True
    
    Exit Function
Catch:
    MsgBox "No se pudo establecer una conexión con el servidor.", vbCritical, "Error de conexión"
End Function

Public Sub TerminateConnection()
    If m_Cn.State = adStateOpen Then
        m_Cn.Close
    End If
    
    Set m_Cn = Nothing
End Sub

'Public Function GetCn() As ADODB.Connection
'    Set GetCn = m_Cn
'End Function

Private Sub Connect()
    If m_Cn.State = adStateClosed Then
        m_Cn.Open
    End If
End Sub

Private Sub Disconnect()
    If Not m_InTrans Then
        If m_Cn.State = adStateOpen Then
            m_Cn.Close
        End If
    End If
End Sub

Public Sub BeginTransaction()
    If Not m_InTrans Then
        m_Cn.Open
        m_Cn.BeginTrans
        m_InTrans = True
    End If
End Sub

Public Sub CommitTransaction()
    If m_InTrans Then
        m_Cn.CommitTrans
        m_Cn.Close
        m_InTrans = False
    End If
End Sub

Public Sub RollbackTransaction()
    If m_InTrans Then
        m_Cn.RollbackTrans
        m_Cn.Close
        m_InTrans = False
    End If
End Sub

Public Function GetRs(ByVal Query As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Connect
    
    Set rs = New ADODB.Recordset
    
    If Not m_InTrans Then
        rs.Open Query, m_Cn, adOpenStatic, adLockReadOnly
        Set rs.ActiveConnection = Nothing
    Else
        rs.Open Query, m_Cn, adOpenDynamic, adLockOptimistic
    End If
    
    Disconnect
    
    Set GetRs = rs
End Function

Public Function GetTable(ByVal Table As String, _
                         Optional ByVal Where As String, _
                         Optional ByVal Fields As String = "*") As ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & Fields & " FROM " & Table
    
    If Len(Where) Then
        sql = sql & " WHERE " & Where
    End If
    
    Set GetTable = GetRs(sql)
End Function

Public Function ConnectRS(ByRef rs As ADODB.Recordset)
    Set rs.ActiveConnection = m_Cn
End Function

Public Sub CloseRS(ByRef rs As ADODB.Recordset, _
                   Optional ByVal Destroy As Boolean = True)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
        
        If Destroy Then
            Set rs = Nothing
        End If
    End If
End Sub

Public Function EmptyRS(ByRef rs As ADODB.Recordset) As Boolean
'On Error Resume Next
'    EmptyRS = True
    EmptyRS = (rs.BOF And rs.EOF)
End Function

Public Function OptimizeField(ByVal rs As Recordset, ByVal FieldName As String) As Boolean
On Error Resume Next
    ' Only work in Client Side.
    If rs.CursorLocation = adUseClient Then
        rs.Fields(FieldName).Properties("OPTIMIZE").Value = True
        OptimizeField = True
    End If
End Function

Public Sub ExecuteQuery(ByRef Query As String, _
                        Optional ByRef RecordsAffected As Long)
    Connect
    m_Cn.Execute Query, RecordsAffected, adCmdText + adExecuteNoRecords
    Disconnect
End Sub

Public Function ExecuteDelete(ByVal Table As String, ByVal Where As String)
    Connect
    m_Cn.Execute "DELETE FROM " & Table & " WHERE " & Where, , adCmdText + adExecuteNoRecords
    Disconnect
End Function

Public Function ExecuteScalar(ByVal Query As String, _
                              Optional ByVal ValueIfNull) As Variant
    Dim rs  As ADODB.Recordset
    Dim RetVal As Variant
    
    If IsMissing(ValueIfNull) Then
        RetVal = vbNullString
    Else
        RetVal = ValueIfNull
    End If
    
    Set rs = GetRs(Query)
    
    If Not (rs.BOF And rs.EOF) Then
        If Not IsNull(rs.Collect(0)) Then
            RetVal = rs.Collect(0)
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    ExecuteScalar = RetVal
End Function

Public Function GetData(ByVal Table As String, _
                        ByVal Field As String, _
                        ByVal Where As String, _
                        Optional ByVal ValueIfNull) As Variant
    GetData = ExecuteScalar("SELECT " & Field & " FROM " & Table & " WHERE " & Where, ValueIfNull)
End Function

Public Function GetNext(ByVal Table As String, _
                        ByVal Field As String, _
                        Optional ByVal Where As String = vbNullString, _
                        Optional ByVal ValueIfNull As Long = 1) As Long
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim RetVal As Long
    
    sql = "SELECT MAX(" & Field & ")+1 FROM " & Table

    If Len(Where) Then
        sql = sql & " WHERE " & Where
    End If
        
    Set rs = GetRs(sql)
    
    If Not (rs.BOF And rs.EOF) Then
        If Not IsNull(rs.Collect(0)) Then
            RetVal = rs.Collect(0)
        Else
            RetVal = ValueIfNull
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetNext = RetVal
End Function

Public Function GetCount(ByVal Table As String, _
                         Optional ByVal Where As String = vbNullString) As Long
    Dim sql As String
    Dim rs  As ADODB.Recordset
    Dim RetVal As Long
    
    sql = "SELECT COUNT(*) FROM " & Table
    
    If Len(Where) Then
        sql = sql & " WHERE " & Where
    End If
    
    Set rs = GetRs(sql)
    
    If Not (rs.BOF And rs.EOF) Then
        If Not IsNull(rs.Collect(0)) Then
            RetVal = rs.Collect(0)
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetCount = RetVal
End Function

Public Function GetLastId() As Long
    'GetLastId = ExecuteScalar("SELECT @@Identity") 'MYSQL
    GetLastId = ExecuteScalar("SELECT last_insert_rowid()") 'SQLITE
End Function

Public Function SQLText(ByVal Value As String) As String
    If Len(Value) Then
        SQLText = "'" & Replace(Value, "'", "''") & "'"
    Else
        SQLText = "NULL"
    End If
End Function

Public Function SQLDate(ByVal Value As Date) As String
    SQLDate = "'" & Format$(Value, "yyyy-mm-dd") & "'"
End Function

Public Function SQLDateTime(ByVal Value As Date) As String
    SQLDateTime = "'" & Format$(Value, "yyyy-mm-dd hh:nn:ss") & "'"
End Function

Public Function SQLTime(ByVal Value As Date) As String
    SQLTime = "'" & Format$(Value, "hh:nn:ss") & "'"
End Function

Public Function SQLNum(ByVal Value As Double) As String
    SQLNum = Trim$(str$(Value))
End Function

Public Function SQLBool(ByVal Value As Boolean) As String
    SQLBool = IIf(Value, "TRUE", "FALSE")
End Function

Public Function Nz(ByVal Value As Variant, _
                   Optional ByVal ValueIfNull As String = vbNullString) As Variant
    If IsNull(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

Public Function GetFieldType(ByRef fld As ADODB.Field) As EFieldType
    Select Case fld.Type
        Case adBigInt, adInteger, adNumeric, _
             adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, _
             adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric
            
            GetFieldType = fdtLong
        
        Case adDecimal, adDouble, adSingle
            GetFieldType = fdtDecimal
            
        Case adCurrency
            GetFieldType = fdtCurrency
        
        Case adVarWChar, adLongVarBinary, adLongVarChar, adLongVarWChar, _
            adVarChar, adWChar, adChar, adChapter, adBSTR
            
            GetFieldType = fdtString
            
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
        
            GetFieldType = fdtDate
            
        Case adBoolean
        
            GetFieldType = fdtBoolean
            
        Case Else
        
            GetFieldType = fdtOther
    End Select
End Function

Public Sub FillCombo(ByRef cbo As ComboBox, _
                     ByVal Table As String, _
                     ByVal DisplayField As String, _
                     ByVal KeyField As String, _
                     Optional ByVal Where As String = vbNullString, _
                     Optional ByVal FirstBlank As Boolean = False, _
                     Optional ByVal SelectFirst As Boolean = False)
                     
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT " & DisplayField & "," & KeyField
    sql = sql & " FROM " & Table
    
    If Len(Where) Then
        sql = sql & " WHERE " & Where
    End If
    
    sql = sql & " ORDER BY " & KeyField
        
    Set rs = GetRs(sql)
            
    With cbo
        .Clear
        
        If FirstBlank Then
            .AddItem vbNullString
        End If
        
        Do While Not rs.EOF
            .AddItem rs.Collect(DisplayField)
            .ItemData(.NewIndex) = rs.Collect(KeyField)
            rs.MoveNext
        Loop
        
        If SelectFirst Then
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
        End If
    End With

    rs.Close
    Set rs = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private m_String As String
Private m_ChunkSize As Long
Private m_Pos As Long
Private m_Len As Long

Public Property Get Length() As Long
    Length = m_Pos \ 2
End Property

Public Property Get Capacity() As Long
    Capacity = m_Len \ 2
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = m_ChunkSize \ 2
End Property

Public Property Let ChunkSize(ByVal iChunkSize As Long)
    ' Set the chunksize.  We multiply by 2 because internally
    ' we are considering bytes:
    m_ChunkSize = iChunkSize * 2
End Property

Public Property Get ToString() As String
    If m_Pos > 0 Then
        ToString = Left$(m_String, m_Pos \ 2)
    End If
End Property

Public Sub SetString(ByRef str As String)
    Dim lLen As Long
   
    lLen = LenB(str)
    If lLen = 0 Then
        Clear
    Else
        If m_Len < lLen Then
            ' Need to expand string to accommodate:
            Do
                m_String = m_String & Space$(m_ChunkSize \ 2)
                m_Len = m_Len + m_ChunkSize
            Loop While m_Len < lLen
        End If
        
        CopyMemory ByVal StrPtr(m_String), ByVal StrPtr(str), lLen
        m_Pos = lLen
    End If
End Sub

Public Sub Append(ByRef str As String)
    Dim lLen As Long

    lLen = LenB(str)
    
    If (m_Pos + lLen) > m_Len Then
        m_String = m_String & Space$(m_ChunkSize \ 2)
        m_Len = m_Len + m_ChunkSize
    End If
    
    CopyMemory ByVal UnsignedAdd(StrPtr(m_String), m_Pos), ByVal StrPtr(str), lLen
    m_Pos = m_Pos + lLen
End Sub

Public Sub Insert(ByVal Index As Long, ByRef str As String)
    Dim lLen As Long
    Dim lPos As Long
    Dim lSize As Long
   
    ' is Index within bounds?
    If (Index * 2 > m_Pos) Then
        Err.Raise 9
    Else
        lLen = LenB(str)
        
        If (m_Pos + lLen) > m_Len Then
            m_String = m_String & Space$(m_ChunkSize \ 2)
            m_Len = m_Len + m_ChunkSize
        End If
          
        ' Move existing characters from current position
        lPos = UnsignedAdd(StrPtr(m_String), Index * 2)
        lSize = m_Pos - Index * 2
          
        ' moving from Index to Index + lLen
        CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
          
        ' Insert new characters:
        CopyMemory ByVal lPos, ByVal StrPtr(str), lLen
          
        m_Pos = m_Pos + lLen
    End If
End Sub

Public Sub Remove(ByVal Index As Long, ByVal lLen As Long)
    Dim lSrc As Long
    Dim lDst As Long
    Dim lSize As Long

    ' is Index within bounds?
    If (Index * 2 > m_Pos) Then
        Err.Raise 9
    Else
        ' is there sufficient length?
        If ((Index + lLen) * 2 > m_Pos) Then
            Err.Raise 9
        Else
            ' Need to copy characters from Index*2 to m_Pos back by lLen chars:
            lSrc = UnsignedAdd(StrPtr(m_String), (Index + lLen) * 2)
            lDst = UnsignedAdd(StrPtr(m_String), Index * 2)
            lSize = (m_Pos - (Index + lLen) * 2)
            CopyMemory ByVal lDst, ByVal lSrc, lSize
            m_Pos = m_Pos - lLen * 2
        End If
    End If
End Sub

Public Function Find(ByVal sToFind As String, _
                     Optional ByVal lStartIndex As Long = 1, _
                     Optional ByVal compare As VbCompareMethod = vbTextCompare _
                     ) As Long
    Dim lInstr As Long
   
    If (lStartIndex > 0) Then
        lInstr = InStr(lStartIndex, m_String, sToFind, compare)
    Else
        lInstr = InStr(m_String, sToFind, compare)
    End If
    
    If (lInstr < m_Pos \ 2) Then
        Find = lInstr
    End If
End Function

Public Sub HeapMinimize()
    Dim iLen As Long
   
    ' Reduce the string size so only the minimal chunks
    ' are allocated:
    If (m_Len - m_Pos) > m_ChunkSize Then
        iLen = m_Len
        Do While (iLen - m_Pos) > m_ChunkSize
            iLen = iLen - m_ChunkSize
        Loop
        m_String = Left$(m_String, iLen \ 2)
        m_Len = iLen
    End If
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
    ' This function is useful when doing pointer arithmetic,
    ' but note it only works for positive values of Incr

    If Start And &H80000000 Then 'Start < 0
        UnsignedAdd = Start + Incr
    ElseIf (Start Or &H80000000) < -Incr Then
        UnsignedAdd = Start + Incr
    Else
        UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
    End If
End Function

Private Sub Class_Initialize()
    ' The default allocation: 8192 characters.
    m_ChunkSize = 16384
End Sub

' METODOS AGREGADOS

Public Sub Clear()
    m_String = ""
    m_Pos = 0
    m_Len = 0
End Sub

Public Sub Replace(Find As String, Replace As String)
    SetString VBA.Replace(Left$(m_String, m_Pos \ 2), Find, Replace)
End Sub

Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Public Function ReadINI(ByVal SectionName As String, _
                        ByVal KeyName As String, _
                        Optional ByVal DefaultValue As String = vbNullString, _
                        Optional ByVal INIPath As String = vbNullString) As String
                        
    Dim lReturn   As Long
    Dim sBuffer   As String * 256
    
    If Len(INIPath) = 0 Then
        INIPath = gPathINI
    End If
    
    lReturn = GetPrivateProfileString(LCase$(SectionName), LCase$(KeyName), DefaultValue, sBuffer, Len(sBuffer), INIPath)
    
    If lReturn Then
        ReadINI = Left$(sBuffer, lReturn)
    Else
        ReadINI = DefaultValue
        WriteINI SectionName, KeyName, DefaultValue, INIPath
    End If
End Function

Public Function WriteINI(ByVal SectionName As String, _
                         ByVal KeyName As String, _
                         ByVal Value As String, _
                         Optional ByVal INIPath As String = vbNullString) As Long
                         
    If Len(INIPath) = 0 Then
        INIPath = gPathINI
    End If
    
    WriteINI = WritePrivateProfileString(LCase$(SectionName), LCase$(KeyName), Value, INIPath)
End Function

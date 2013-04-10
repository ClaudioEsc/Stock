Attribute VB_Name = "modUsuario"
Option Explicit

Public Function Login(ByVal UserName As String, _
                      ByVal Password As String) As Boolean
    Dim rs As ADODB.Recordset

    With New CString
        .Append " SELECT usuario, es_admin"
        .Append " FROM usuarios"
        .Append " WHERE usuario = " & SQLText(UserName)
        .Append " AND contrasena = " & SQLText(Password)

        Set rs = GetRs(.ToString)
    End With

    If Not EmptyRS(rs) Then
        Login = True
    End If
End Function



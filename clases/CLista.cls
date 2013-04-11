VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Titulo   As String
Public Tabla    As String
Public Campos   As Collection

Public Sub AgregarCampo(ByVal Nombre As String, _
                        ByVal Titulo As String, _
                        Optional ByVal Alineacion As ECampoAlineacion = ECampoAlineacion.caIzquierda, _
                        Optional ByVal Tipo As ECampoTipo = ECampoTipo.ctTexto)
    Dim c As CListaCampo
    
    Set c = New CListaCampo
    
    With c
        .Nombre = Nombre
        .Titulo = Titulo
        .Alineacion = Alineacion
        .Tipo = Tipo
    End With
    
    Campos.Add c
    
    Set c = Nothing
End Sub

Public Function GetSQL() As String
    Dim cs  As String
    Dim c   As CListaCampo
    
    For Each c In Campos
        cs = cs & c.Nombre & ","
    Next
    
    GetSQL = "SELECT " & Left$(cs, Len(cs) - 1) & " FROM " & Tabla
End Function

Private Sub Class_Initialize()
    Set Campos = New Collection
End Sub

Private Sub Class_Terminate()
    Set Campos = Nothing
End Sub


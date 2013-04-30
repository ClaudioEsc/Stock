VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFormABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Id() As String
End Property

Public Function ShowModal(Optional ByRef OwnerForm As Form) As EModalResult
End Function

Public Sub Iniciar(ByVal IsNew As Boolean, Optional ByVal Id As String)
End Sub

Public Function Eliminar(ByVal Id As String) As Boolean
End Function

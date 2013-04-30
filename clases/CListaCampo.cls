VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ECampoAlineacion
    caIzquierda
    caDerecha
    caCentro
End Enum

Public Enum ECampoTipo
    ctNumero
    ctTexto
    ctFecha
    ctHora
    ctFechaHora
    ctBool
End Enum

Public Indice       As Long
Public Nombre       As String
Public Titulo       As String
Public Alineacion   As ECampoAlineacion
Public Tipo         As ECampoTipo
Public PermiteBuscar As Boolean
Public Formato      As String

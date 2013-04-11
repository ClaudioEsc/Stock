VERSION 5.00
Begin VB.Form frmConexion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de la conexión"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConexion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   315
      Left            =   5340
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtPathDB 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Base de datos:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    WriteINI "db", "path", txtPathDB.Text
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    With New CCommonDialog
        .Filter = "Base de datos SQLite|*.s3db"
        
        If .ShowOpen Then
            txtPathDB.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


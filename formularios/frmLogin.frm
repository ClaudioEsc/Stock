VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iniciar sesión"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   480
      Width           =   2595
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2595
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ModalResult As VbMsgBoxResult

Public Function ShowModal() As VbMsgBoxResult
    txtUsuario.Text = ReadINI("DB", "User")
    Me.Show vbModal
    ShowModal = m_ModalResult
End Function

Private Sub cmdAceptar_Click()
    If Not Login(txtUsuario.Text, txtPassword.Text) Then
        MsgBox "Usuario y/o contraseña incorrectos.", vbCritical
    Else
        m_ModalResult = vbOK
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    m_ModalResult = vbCancel
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Len(txtUsuario.Text) <> 0 Then
        txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteINI "DB", "User", txtUsuario.Text
End Sub

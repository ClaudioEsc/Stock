VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   780
      Width           =   1035
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2520
      Top             =   60
   End
   Begin VB.PictureBox pic 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   3015
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3015
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "versión"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.00"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub IniciarPublicidad()
    cmdSalir.Visible = False
    tmr.Enabled = True
    Me.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = Format$(App.Major, "00") & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
End Sub

Private Sub tmr_Timer()
    Unload Me
End Sub

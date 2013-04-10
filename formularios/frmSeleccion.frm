VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSeleccion 
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSeleccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFiltro 
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5700
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6900
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cboCampo 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5318
      _Version        =   393216
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ETipoSeleccion
    tsClientes
    tsProductos
End Enum

Private m_TipoSeleccion As ETipoSeleccion
Private m_OK As Boolean
Private m_Rs As ADODB.Recordset

Public Property Get ItemSeleccionado(Optional ByVal CampoRetorno = 0) As String
    ItemSeleccionado = m_Rs.Collect(0)
End Property

Private Sub cmdOK_Click()
    If cmdOk.Enabled Then
        m_Rs.Move grd.Row - 1, 1
        m_OK = True
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Mostrar()

On Error GoTo ErrHandler

    If Len(txtFiltro.Text) = 0 Then
        m_Rs.Filter = vbNullString
    Else
        If cboCampo.ListIndex = -1 Then
            cboCampo.ListIndex = 0
        End If
        
         m_Rs.Filter = m_Rs(cboCampo.ListIndex).Name & " LIKE " & SQLText("%" & txtFiltro.Text & "%")
    End If
    
    FillGrid grd, m_Rs
    
Finally:
'    grd.AutoSize
'    grd.SelectRow 1
    cmdOk.Enabled = Not EmptyRS(m_Rs)
    
    Exit Sub
ErrHandler:
    ErrorReport "frmSeleccion", "Mostrar"
    Resume Finally
End Sub

Public Function Seleccionar(ByVal TipoSeleccion As ETipoSeleccion) As Boolean
    Dim Consulta As String
    Dim Titulos() As Variant
    Dim CampoDefecto As Long
    Dim i As Long
    
    m_OK = False
    m_TipoSeleccion = TipoSeleccion
    
    Select Case m_TipoSeleccion
        Case tsClientes
            Me.Caption = "Seleccionar cliente"
            Consulta = "SELECT id, nombre FROM clientes ORDER BY id"
            Titulos = Array("Id", "Nombre")
            CampoDefecto = 1
            
        Case tsProductos
            Me.Caption = "Seleccionar producto"
            Consulta = "SELECT id, descripcion, familia, marca FROM v_productos ORDER BY id"
            Titulos = Array("Id", "Referencia", "Descripción", "Familia", "Marca")
            CampoDefecto = 2
    End Select
        
    Set m_Rs = GetRs(Consulta)

    With grd
        .Redraw = False
        .Cols = UBound(Titulos) + .FixedCols + 1
        .Rows = .FixedRows + 1
    
        For i = LBound(Titulos) To UBound(Titulos)
            cboCampo.AddItem Trim$(Titulos(i))
            grd.TextMatrix(0, i + grd.FixedCols) = Trim$(Titulos(i))
        Next
        
        grd.Redraw = True
    End With
    
    If cboCampo.ListCount > 0 Then
       cboCampo.ListIndex = CampoDefecto
    End If
        
    Mostrar
    
    Me.Show vbModal
    Seleccionar = m_OK
End Function

Private Sub grd_DblClick()
    If grd.MouseRow = 0 Then
'        grd.AutoSize grd.MouseCol
    Else
        cmdOK_Click
    End If
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_Terminate()
    CloseRS m_Rs
End Sub

Private Sub txtFiltro_GotFocus()
    HLText txtFiltro
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        grd.SetFocus
    End If
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Mostrar
        
        If grd.Rows > 1 Then
            grd.SetFocus
        Else
            txtFiltro.SetFocus
        End If
    End If
End Sub

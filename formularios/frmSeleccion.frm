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
   Begin VB.CommandButton cmdAceptar 
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
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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

Private m_Seleccion     As String
Private m_Rs            As ADODB.Recordset
Private m_ModalResult   As EModalResult
Private m_CampoNombre   As String
Private m_CampoTipo     As EFieldType
Private m_Anchor        As CAnchor

Public Property Get ModalResult() As EModalResult
    ModalResult = m_ModalResult
End Property

Public Property Get ItemSeleccionado() As String
    ItemSeleccionado = m_Seleccion
End Property

Public Sub Iniciar(ByVal sql As String, _
                   ByVal CamposTitulos As String, _
                   Optional ByVal Titulo As String = "Seleccionar", _
                   Optional ByVal CampoBusqueda As Long = 1)

    Dim i As Long
    Dim Titulos() As String
    
On Error GoTo Catch
    Set m_Rs = GetRs(sql)
        
    If EmptyRS(m_Rs) Then
        MsgBox "No hay datos para mostrar.", vbExclamation, gAppName
    Else
        With grd
            .Clear
            .Rows = 1
            .Cols = m_Rs.Fields.Count
        End With
        
        cboCampo.Clear
        Titulos = Split(CamposTitulos, ",")
        
        For i = 0 To m_Rs.Fields.Count - 1
            grd.TextMatrix(0, i) = Trim$(Titulos(i))
            cboCampo.AddItem Trim$(Titulos(i))

            Select Case GetFieldType(m_Rs.Fields(i))
                Case fdtLong, fdtDecimal, fdtCurrency, fdtDate
                    grd.ColAlignment(i) = flexAlignRightCenter
                    
                Case fdtBoolean
                    grd.ColAlignment(i) = flexAlignCenterCenter
                    
                Case Else
                    grd.ColAlignment(i) = flexAlignLeftCenter
            End Select
        Next
        
        cboCampo.ListIndex = CampoBusqueda
        Me.Caption = Titulo
        Me.Show vbModal
    End If
    
    Exit Sub
Catch:
    ErrorReport "Iniciar", "frmSeleccion"
    Unload Me
End Sub

Private Sub Mostrar()
    Dim Filtro      As String
    Dim Campo       As String
    Dim CampoTipo   As EFieldType
    Dim i           As Long
    Dim j           As Long
    
    Screen.MousePointer = vbHourglass
    Filtro = Trim$(txtFiltro.Text)

On Error GoTo Catch
    If Len(Filtro) = 0 Then
        m_Rs.Filter = vbNullString
    Else
        Select Case m_CampoTipo
            Case fdtLong, fdtDecimal, fdtCurrency
                m_Rs.Filter = m_CampoNombre & " = " & ToNumber(Filtro)
                    
            Case fdtDate
                If IsDate(Filtro) Then
                    m_Rs.Filter = m_CampoNombre & " = " & CDate(Filtro)
                End If
            
            Case Else
                m_Rs.Filter = m_CampoNombre & " LIKE '%" & Filtro & "%'"
        End Select
    End If
    
    With grd
        .Redraw = False
        .Rows = 1
    
        If EmptyRS(m_Rs) Then
            cmdAceptar.Enabled = False
        Else
            cmdAceptar.Enabled = True
            
            .Rows = m_Rs.RecordCount + 1
            
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = m_Rs.Fields(j) & vbNullString
                Next
                
                m_Rs.MoveNext
            Next
        End If
        
        .Redraw = True
    End With

On Error Resume Next
    GridAutoSize grd
    GridSelectRow grd, 1
    Screen.MousePointer = vbDefault

    Exit Sub
Catch:
    ErrorReport "frmSeleccion", "Mostrar"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboCampo_Click()
    m_CampoNombre = m_Rs.Fields(cboCampo.ListIndex).Name
    m_CampoTipo = GetFieldType(m_Rs.Fields(cboCampo.ListIndex))
    Mostrar
End Sub

Private Sub cmdCancelar_Click()
    m_ModalResult = mrCancel
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    If grd.Rows > 1 And grd.Row > 0 Then
        m_Seleccion = grd.TextMatrix(grd.Row, 0)
        m_ModalResult = mrOK
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set m_Anchor = New CAnchor
    With m_Anchor
        .AddControl txtFiltro, apLeft + apRight
        .AddControl grd, apAll
        .AddControl cmdAceptar, apBottom + apRight
        .AddControl cmdCancelar, apBottom + apRight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseRS m_Rs
    Set m_Anchor = Nothing
End Sub

Private Sub grd_DblClick()
    If cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cmdAceptar.Enabled Then
            cmdAceptar_Click
        End If
    End If
End Sub

Private Sub txtFiltro_GotFocus()
    HLText txtFiltro
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Mostrar
    End If
    
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyDown) And grd.Rows > 1 Then
        grd.SetFocus
    End If
End Sub

'Constructores por defecto
Public Function IniciarProductos() As Boolean
    Dim sql As New CString
    
    With sql
        .Append " SELECT codigo, descripcion"
        .Append " FROM productos"
        .Append " ORDER BY codigo"
                
        Iniciar .ToString, "Código, Descripción"
    End With
End Function

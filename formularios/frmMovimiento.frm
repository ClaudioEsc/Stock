VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMovimiento 
   Caption         =   "Movimiento de stock"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimiento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNumero 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   7740
      ScaleHeight     =   735
      ScaleWidth      =   1695
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   1695
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.OptionButton optEgreso 
      Caption         =   "Egreso"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   1035
   End
   Begin VB.OptionButton optIngreso 
      Caption         =   "Ingreso"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.TextBox txtDescripcion 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1140
      Width           =   5835
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "+ Agregar"
      Height          =   315
      Left            =   8460
      TabIndex        =   3
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscarProducto 
      Caption         =   "..."
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox txtIdProducto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   6480
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7260
      TabIndex        =   5
      Top             =   6480
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4275
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7541
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   900
      Width           =   810
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9180
      TabIndex        =   12
      Top             =   5940
      Width           =   150
   End
   Begin VB.Shape shpTotal 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   495
      Left            =   120
      Top             =   5820
      Width           =   9315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   7440
      TabIndex        =   10
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   495
   End
End
Attribute VB_Name = "frmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_IDPRODUCTO    As Long = 0
Private Const COL_DESCRIPCION   As Long = 1
Private Const COL_CANTIDAD      As Long = 2
Private Const COL_PRECIO        As Long = 3
Private Const COL_IMPORTE       As Long = 4
Private Const COL_COUNT         As Long = 5

Private Anchor          As CAnchor
Private m_Precio        As Currency
Private m_Id            As Long
Private m_ModalResult   As EModalResult
Private m_IsNew         As Boolean

Public Property Get Id() As Long
    Id = m_Id
End Property

Public Function ShowModal(Optional ByRef OwnerForm As Form) As EModalResult
    Me.Show vbModal, OwnerForm
    ShowModal = m_ModalResult
End Function

Public Sub Iniciar(ByVal IsNew As Boolean, _
                   Optional ByVal Id As String = vbNullString)
    m_IsNew = IsNew
    m_Id = Val(Id)
End Sub

Private Sub Mostrar()
    Dim rs          As ADODB.Recordset
    Dim rsDetalle   As ADODB.Recordset
    
On Error GoTo ErrorHandler
    Set rs = GetTable("movimientos", "id = " & SQLNum(m_Id))
    Set rsDetalle = GetTable("vw_movimientos_det", "idmovimiento = " & SQLNum(m_Id))
    
    If EmptyRS(rs) Then
        MsgBox "No se encontró el movimiento.", vbExclamation
    Else
        With rs
            If .Collect("tipo") = MOVIMIENTO_ENTRADA Then
                optIngreso.Value = True
            Else
                optEgreso.Value = True
            End If
            
            txtId.Text = Format$(m_Id)
            txtFecha.Text = .Collect("fecha")
        End With
        
        With rsDetalle
            Do While Not .EOF
                AgregarDetalle .Collect("idproducto"), _
                               .Collect("descripcion"), _
                               .Collect("cantidad"), _
                               .Collect("precio"), _
                               .Collect("iddetalle")
                .MoveNext
            Loop
        End With
    End If
        
Finally:
    CloseRS rs
    CloseRS rsDetalle
    
    Exit Sub
ErrorHandler:
    ErrorReport "frmMovimientos", "Mostrar"
    Resume Finally
End Sub

Public Function Eliminar(ByVal Id As Long) As Boolean
On Error GoTo ErrorHandler
    ExecuteDelete "movimientos", "id = " & SQLNum(Id)
    Eliminar = True

    Exit Function
ErrorHandler:
    ErrorReport "frmMovimientos", "Eliminar"
End Function

Private Function Validar() As Boolean
    If grd.Rows = 1 Then
        MsgBox "Debe ingresar al menos un producto.", vbExclamation
        txtIdProducto.SetFocus
        Exit Function
    End If

    Validar = True
End Function

Private Function Guardar() As Boolean
    If Not Validar() Then
        Exit Function
    End If
    
On Error GoTo ErrorHandler
    BeginTransaction
    
    If m_IsNew Then
        With New CString
            .Clear
            .Append "INSERT INTO movimientos"
            .Append "(fecha, tipo)"
            .Append "VALUES"
            .Append "(" & SQLDate(CDate(txtFecha.Text))
            .Append "," & SQLText(IIf(optIngreso.Value, MOVIMIENTO_ENTRADA, MOVIMIENTO_SALIDA))
            .Append ")"
            
            ExecuteQuery .ToString
        End With
        
        m_Id = GetLastId()
        
        GuardarDetalle
    End If
    
    CommitTransaction
    Guardar = True
    
    Exit Function
ErrorHandler:
    ErrorReport "frmProducto", "Guardar"
    RollbackTransaction
End Function

Public Sub GuardarDetalle()
    Dim IdProducto  As Long
    Dim Cantidad    As Double
    Dim Precio      As Currency
    Dim sql         As CString
    Dim Tipo        As Long
    Dim i           As Long
    
    Set sql = New CString
    Tipo = IIf(optIngreso.Value, 1, -1)
    
    For i = 1 To grd.Rows - 1
        IdProducto = Val(grd.TextMatrix(i, COL_IDPRODUCTO))
        Cantidad = CDbl(grd.TextMatrix(i, COL_CANTIDAD))
        Precio = CDbl(grd.TextMatrix(i, COL_PRECIO))
        
        With sql
            .Clear
            .Append "INSERT INTO movimientos_det"
            .Append "(idmovimiento, idproducto, cantidad, precio)"
            .Append "VALUES"
            .Append "(" & SQLNum(m_Id)
            .Append "," & SQLNum(IdProducto)
            .Append "," & SQLNum(Cantidad)
            .Append "," & SQLNum(Precio)
            .Append ")"
            
            ExecuteQuery .ToString
            
            .Clear
            .Append " UPDATE productos SET"
            .Append " stock = stock + " & SQLNum(Cantidad * Tipo)
            .Append " WHERE id = " & SQLNum(IdProducto)
            
            ExecuteQuery .ToString
        End With
    Next
End Sub

Private Sub cmdAceptar_Click()
    If Guardar Then
        m_ModalResult = mrOK
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    m_ModalResult = mrCancel
    Unload Me
End Sub

Private Sub cmdAgregar_Click()
    If ValidarDetalle Then
        AgregarDetalle Val(txtIdProducto.Text), txtDescripcion.Text, ToNumber(txtCantidad.Text), m_Precio
        LimpiarDetalle
        txtIdProducto.SetFocus
    End If
End Sub

Private Sub cmdBuscarProducto_Click()
    With New frmSeleccion
        .IniciarProductos
        
        If .ModalResult = mrOK Then
            If CargarProducto(Val(.ItemSeleccionado)) Then
                txtCantidad.SetFocus
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    Set Anchor = New CAnchor
    With Anchor
        .AddControl grd, apAll
        .AddControl shpTotal, apBottom + apLeft + apRight
        .AddControl lblTotal, apBottom + apRight
        .AddControl cmdAceptar, apBottom + apRight
        .AddControl cmdCancelar, apBottom + apRight
        .AddControl picNumero, apRight
    End With
    
    With grd
        .Rows = 1
        .Cols = COL_COUNT
        
        GridInitCol grd, COL_IDPRODUCTO, "Código", 1000, gcaRight
        GridInitCol grd, COL_DESCRIPCION, "Descripción", 2500
        GridInitCol grd, COL_CANTIDAD, "Cantidad", 1000, gcaRight
        GridInitCol grd, COL_PRECIO, "Precio", 1000, gcaRight
        GridInitCol grd, COL_IMPORTE, "Importe", 1000, gcaRight
    End With
    
    txtFecha.Text = FormatDateTime(Date)
    LimpiarDetalle
    
    If Not m_IsNew And m_Id > 0 Then
        Mostrar
        HabilitarEdicion False
    End If
End Sub

Private Function CargarProducto(ByVal IdProducto As Long) As Boolean
    Dim rs As ADODB.Recordset
    
On Error GoTo Catch
    Set rs = GetTable("productos", "id = " & SQLNum(IdProducto), "descripcion, precio")
    
    If EmptyRS(rs) Then
        MsgBox "No se encontró el producto.", vbExclamation
        LimpiarDetalle
    Else
        With rs
            txtIdProducto.Text = Format$(IdProducto)
            txtDescripcion.Text = .Collect("descripcion")
            m_Precio = .Collect("precio")
        End With
        
        CargarProducto = True
    End If
    
Finally:
    CloseRS rs
    
    Exit Function
Catch:
    ErrorReport "frmMovimiento", "CargarProducto"
    Resume Finally
End Function

Private Sub LimpiarDetalle()
    txtIdProducto.Text = vbNullString
    txtDescripcion.Text = vbNullString
    txtCantidad.Text = Format$(1, "#0.00")
    m_Precio = 0
End Sub

Private Sub grd_DblClick()
    EliminarDetalle
End Sub

Private Sub grd_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete:   EliminarDetalle
    End Select
End Sub

Private Sub txtCantidad_GotFocus()
    HLText txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAgregar.Value = True
    End If
End Sub

Private Sub txtCantidad_LostFocus()
    txtCantidad.Text = Format$(ToNumber(txtCantidad.Text), "#0.00")
End Sub

Private Sub txtIdProducto_GotFocus()
    HLText txtIdProducto
End Sub

Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtIdProducto.Text) = 0 Then
            cmdBuscarProducto.Value = True
        Else
            If CargarProducto(Val(txtIdProducto.Text)) Then
                txtCantidad.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtIdProducto_Validate(Cancel As Boolean)
    If Len(txtIdProducto.Text) = 0 Then
        LimpiarDetalle
    Else
        If Not CargarProducto(Val(txtIdProducto.Text)) Then
            Cancel = True
        End If
    End If
End Sub

Private Sub AgregarDetalle(ByVal IdProducto As Long, _
                           ByVal Descripcion As String, _
                           ByVal Cantidad As Double, _
                           ByVal Precio As Currency, _
                           Optional ByVal IdDetalle As Long = 0)
    Dim Row As Long

    Row = GridFindRow(grd, Format$(IdProducto), 1, COL_IDPRODUCTO)

    With grd
        ' Si el producto no se encuentra en la lista, lo agrego
        If Row = -1 Then
            .AddItem vbNullString
            Row = .Rows - 1

            .RowData(Row) = IdDetalle
            .TextMatrix(Row, COL_IDPRODUCTO) = Format$(IdProducto)
            .TextMatrix(Row, COL_DESCRIPCION) = Descripcion
        Else
        'Si el producto ya esta en la lista, solamente sumo la cantidad
            Cantidad = Cantidad + CDbl(.TextMatrix(Row, COL_CANTIDAD))
            Precio = CDbl(.TextMatrix(Row, COL_PRECIO))
        End If

        .TextMatrix(Row, COL_CANTIDAD) = FormatNumber(Cantidad, 2)
        .TextMatrix(Row, COL_PRECIO) = FormatCurrency(Precio, 2)
        .TextMatrix(Row, COL_IMPORTE) = FormatCurrency(Cantidad * Precio, 2)
    End With

    CalcularTotal
End Sub

Private Function ValidarDetalle() As Boolean
    If Len(txtIdProducto.Text) = 0 Then
        MsgBox "Seleccione un producto.", vbExclamation
        txtIdProducto.SetFocus
        Exit Function
    End If

    If ToNumber(txtCantidad.Text) <= 0 Then
        MsgBox "La cantidad debe ser mayor a 0 (cero).", vbExclamation
        txtCantidad.SetFocus
        Exit Function
    End If

    ValidarDetalle = True
End Function

Private Sub EliminarDetalle()
    If m_IsNew Then
        With grd
            If .Rows > 1 And .Row > 0 Then
                GridRemoveRow grd, .Row
                CalcularTotal
            End If
        End With
    End If
End Sub

Private Sub CalcularTotal()
    Dim i       As Long
    Dim Total   As Currency
    
    For i = 1 To grd.Rows - 1
        Total = Total + CCur(grd.TextMatrix(i, COL_IMPORTE))
    Next
    
    lblTotal.Caption = FormatCurrency(Total, 2)
End Sub

Private Sub HabilitarEdicion(ByVal Valor As Boolean)
    optIngreso.Enabled = Valor
    optEgreso.Enabled = Valor
    txtFecha.Enabled = Valor
    txtIdProducto.Enabled = Valor
    cmdBuscarProducto.Enabled = Valor
    txtCantidad.Enabled = Valor
    cmdAgregar.Enabled = Valor
    cmdAceptar.Visible = Valor
End Sub

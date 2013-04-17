VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMovimiento 
   Caption         =   "Movimiento de stock"
   ClientHeight    =   6735
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
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
      Width           =   1035
   End
   Begin VB.TextBox txtDescripcion 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   900
      Width           =   5835
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "+ Agregar"
      Height          =   315
      Left            =   8460
      TabIndex        =   3
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscarProducto 
      Caption         =   "..."
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox txtCantidad 
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   6240
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7260
      TabIndex        =   5
      Top             =   6240
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4275
      Left            =   120
      TabIndex        =   4
      Top             =   1260
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
      Top             =   660
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
      Top             =   5700
      Width           =   150
   End
   Begin VB.Shape shpTotal 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   495
      Left            =   120
      Top             =   5580
      Width           =   9315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   7440
      TabIndex        =   10
      Top             =   660
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   660
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

Private Anchor      As CAnchor
Private m_Precio    As Double

Private Sub cmdBuscarProducto_Click()
    With New frmSeleccion
        .IniciarProductos
        
        If .ModalResult = mrOk Then
            CargarProducto Val(.ItemSeleccionado)
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
    End With
    
    With grd
        .Rows = 1
        .Cols = COL_COUNT
        
        InitGridCol grd, COL_IDPRODUCTO, "Código", 1000, gcaRight
        InitGridCol grd, COL_DESCRIPCION, "Descripción", 2500, gcaRight
        InitGridCol grd, COL_CANTIDAD, "Cantidad", 1000, gcaRight
        InitGridCol grd, COL_PRECIO, "Precio", 1000, gcaRight
        InitGridCol grd, COL_IMPORTE, "Importe", 1000, gcaRight
    End With
End Sub

Private Sub CargarProducto(ByVal IdProducto As Long)
    Dim rs As ADODB.Recordset
    
On Error GoTo Catch
    Set rs = GetTable("productos", "id = " & SQLNum(IdProducto), "descripcion, precio")
    
    If IsEmpty(rs) Then
        MsgBox "No se encontró el producto.", vbExclamation
        LimpiarProducto
    Else
        With rs
            txtCodigo.Text = Format$(IdProducto)
            txtDescripcion.Text = .Collect("descripcion")
            m_Precio = .Collect("precio")
        End With
    End If
    
Finally:
    CloseRS rs
    
    Exit Sub
Catch:
    ErrorReport "frmMovimiento", "CargarProducto"
    Resume Finally
End Sub

Private Sub LimpiarProducto()
    txtCodigo.Text = vbNullString
    txtDescripcion.Text = vbNullString
    m_Precio = 0
End Sub

Private Sub txtCodigo_Change()
    
End Sub

VERSION 5.00
Begin VB.Form frmProducto 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.TextBox txtStock 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   1035
   End
   Begin VB.TextBox txtStockMinimo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   1035
   End
   Begin VB.TextBox txtPrecio 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   1035
   End
   Begin VB.TextBox txtCosto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.ComboBox cboRubro 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Precio:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Costo:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Stock mínimo:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Stock actual:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   870
   End
End
Attribute VB_Name = "frmProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFormABM

Private m_Id            As Long
Private m_ModalResult   As EModalResult
Private m_IsNew         As Boolean

Public Property Get IFormABM_Id() As String
    IFormABM_Id = Format$(m_Id)
End Property

Public Function IFormABM_ShowModal(Optional ByRef OwnerForm As Form) As EModalResult
    Me.Show vbModal, OwnerForm
    IFormABM_ShowModal = m_ModalResult
End Function

Public Sub IFormABM_Iniciar(ByVal IsNew As Boolean, Optional ByVal Id As String)
    m_IsNew = IsNew
    m_Id = Val(Id)
End Sub

Public Function IFormABM_Eliminar(ByVal Id As String) As Boolean
On Error GoTo ErrorHandler
    ExecuteDelete "productos", "id = " & SQLNum(Val(Id))
    IFormABM_Eliminar = True

    Exit Function
ErrorHandler:
    ErrorReport "frmProducto", "Delete"
End Function

Private Sub Mostrar()
    Dim rs As ADODB.Recordset
    
On Error GoTo ErrorHandler
    Set rs = GetTable("productos", "id = " & SQLNum(m_Id))
        
    If EmptyRS(rs) Then
        MsgBox "No se encontró el producto.", vbExclamation, gAppName
    Else
        With rs
'            txtId.Text = .Collect("id")
            txtCodigo.Text = .Collect("codigo")
            txtDescripcion.Text = Nz(.Collect("descripcion"))
            txtStock.Text = FormatDecimal(.Collect("stock"))
            txtStockMinimo.Text = FormatDecimal(.Collect("stock_minimo"))
            txtPrecio.Text = FormatDecimal(.Collect("precio"))
            txtCosto.Text = FormatDecimal(.Collect("costo"))
            SetItemData cboRubro, .Collect("idrubro")
        End With
    End If
        
Finally:
    CloseRS rs
    
    Exit Sub
ErrorHandler:
    ErrorReport "frmProducto", "ShowData"
    Resume Finally
End Sub

Private Function Validar() As Boolean
    If Len(txtCodigo.Text) = 0 Then
        MsgBox "El código es requerido.", vbExclamation, gAppName
        txtCodigo.SetFocus
        Exit Function
    Else
        If GetCount("productos", "codigo = " & SQLText(txtCodigo.Text)) > 0 Then
            MsgBox "El código ingresado ya fue asignado a otro producto.", vbExclamation
            txtCodigo.SetFocus
            Exit Function
        End If
    End If
    
    If Len(txtDescripcion.Text) = 0 Then
        MsgBox "La descripción es requerida.", vbExclamation, gAppName
        txtDescripcion.SetFocus
        Exit Function
    End If
    
    If cboRubro.ListIndex = -1 Then
        MsgBox "El rubro es requerido.", vbExclamation, gAppName
        cboRubro.SetFocus
        Exit Function
    End If

    Validar = True
End Function

Private Function Guardar() As Boolean
    Dim sql As String
    Dim i As Long
    
    If Not Validar() Then
        Exit Function
    End If
    
On Error GoTo ErrorHandler
    BeginTransaction
    
    If m_IsNew Then
        With New CString
            .Append "INSERT INTO productos"
            .Append "(codigo, descripcion, idrubro, stock_minimo, stock, precio, costo)"
            .Append "VALUES"
            .Append "(" & SQLText(txtCodigo.Text)
            .Append "," & SQLText(txtDescripcion.Text)
            .Append "," & SQLNum(GetItemData(cboRubro))
            .Append "," & SQLNum(ToNumber(txtStockMinimo.Text))
            .Append "," & SQLNum(ToNumber(txtStock.Text))
            .Append "," & SQLNum(ToNumber(txtPrecio.Text))
            .Append "," & SQLNum(ToNumber(txtCosto.Text))
            .Append ")"
            
            ExecuteQuery .ToString
        End With
        
        m_Id = GetLastId()
    Else
        With New CString
            .Append "UPDATE productos SET"
            .Append "  descripcion = " & SQLText(txtDescripcion.Text)
            .Append ", codigo = " & SQLText(txtCodigo.Text)
            .Append ", idrubro = " & SQLNum(GetItemData(cboRubro))
            .Append ", stock_minimo = " & SQLNum(ToNumber(txtStockMinimo.Text))
            .Append ", stock = " & SQLNum(ToNumber(txtStock.Text))
            .Append ", precio = " & SQLNum(ToNumber(txtPrecio.Text))
            .Append ", costo = " & SQLNum(ToNumber(txtCosto.Text))
            .Append " WHERE id = " & SQLNum(m_Id)
            
            ExecuteQuery .ToString
        End With
    End If
    
    CommitTransaction
    Guardar = True
    
    Exit Function
ErrorHandler:
    ErrorReport "frmProducto", "Guardar"
    RollbackTransaction
End Function

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

Private Sub Form_Load()
    FillCombo cboRubro, "rubros", "descripcion", "id"
    txtStockMinimo.Text = FormatDecimal(0)
    txtStock.Text = FormatDecimal(0)
    txtPrecio.Text = FormatDecimal(0)
    txtCosto.Text = FormatDecimal(0)
    
    If Not m_IsNew And m_Id > 0 Then
        Mostrar
    End If
    
    If m_IsNew Then
        Me.Caption = "Producto - Nuevo"
    Else
        Me.Caption = "Producto - Modificando"
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    HLText txtCodigo
End Sub

Private Sub txtCosto_GotFocus()
    HLText txtCosto
End Sub

Private Sub txtCosto_LostFocus()
    txtCosto.Text = Format$(ToNumber(txtCosto.Text), "#0.00")
End Sub

Private Sub txtDescripcion_GotFocus()
    HLText txtDescripcion
End Sub

Private Sub txtPrecio_GotFocus()
    HLText txtPrecio
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.Text = Format$(ToNumber(txtPrecio.Text), "#0.00")
End Sub

Private Sub txtStock_GotFocus()
    HLText txtStock
End Sub

Private Sub txtStock_LostFocus()
    txtStock.Text = Format$(ToNumber(txtStock.Text), "#0.00")
End Sub

Private Sub txtStockMinimo_GotFocus()
    HLText txtStockMinimo
End Sub

Private Sub txtStockMinimo_LostFocus()
    txtStockMinimo.Text = Format$(ToNumber(txtStockMinimo.Text), "#0.00")
End Sub

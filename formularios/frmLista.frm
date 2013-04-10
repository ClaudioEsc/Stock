VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLista 
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   10185
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picNoData 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   2535
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<No se encontraron datos>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   2340
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   1635
      Left            =   60
      TabIndex        =   10
      Top             =   1320
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2884
      _Version        =   393216
      RowHeightMin    =   300
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin ComctlLib.StatusBar sbr 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   7005
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   17463
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10185
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   10185
      Begin VB.CommandButton cmdMenu 
         Caption         =   "A&justar"
         Height          =   435
         Index           =   6
         Left            =   5580
         TabIndex        =   5
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Cerrar"
         Height          =   435
         Index           =   5
         Left            =   6780
         TabIndex        =   6
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Actualizar"
         Height          =   435
         Index           =   4
         Left            =   4500
         TabIndex        =   4
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Exportar"
         Height          =   435
         Index           =   3
         Left            =   3420
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Eliminar"
         Height          =   435
         Index           =   2
         Left            =   2220
         TabIndex        =   2
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Modificar"
         Height          =   435
         Index           =   1
         Left            =   1140
         TabIndex        =   1
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "&Nuevo"
         Height          =   435
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.PictureBox picFiltro 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10185
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   555
      Width           =   10185
      Begin VB.TextBox txtFiltro 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   300
         Width           =   1815
      End
      Begin VB.ComboBox cboCampo 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   1755
      End
      Begin VB.PictureBox picOrden 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   3720
         ScaleHeight     =   555
         ScaleWidth      =   2595
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   2595
         Begin VB.ComboBox cboOrden 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2595
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000010&
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum EAccionABM
    abmNuevo
    abmModificar
    abmEliminar
End Enum

Private Const BTN_NEW       As Long = 0
Private Const BTN_EDIT      As Long = 1
Private Const BTN_DELETE    As Long = 2
Private Const BTN_EXPORT    As Long = 3
Private Const BTN_REFRESH   As Long = 4
Private Const BTN_CLOSE     As Long = 5
Private Const BTN_AUTOSIZE  As Long = 6

Private m_Lista     As EListas
Private m_ObjLista  As CLista

Public Sub Iniciar(ByVal Lista As EListas)
    m_Lista = Lista
End Sub

Public Sub IniciarLista()
    Set m_ObjLista = New CLista
    
    With m_ObjLista
        Select Case m_Lista
            Case lsProductos
                .Titulo = "Productos"
                .Tabla = "productos"
                .AgregarCampo "id", "C�digo"
                .AgregarCampo "descripcion", "Descripci�n"
                .AgregarCampo "stock", "Stock"
            
            Case lsRubros
                .Titulo = "Rubros"
                .Tabla = "rubros"
                .AgregarCampo "id", "C�digo"
                .AgregarCampo "descripcion", "Descripci�n"
                
        End Select
    End With
End Sub

Public Sub IniciarControles()
    Dim i As Long
    Dim TituloCampo As String

    grd.Cols = grd.FixedCols + m_ObjLista.Campos.Count
    
    For i = 1 To m_ObjLista.Campos.Count
        TituloCampo = m_ObjLista.Campos(i).Titulo
        
        cboCampo.AddItem TituloCampo
        
        cboOrden.AddItem TituloCampo & " [ASC]"
        cboOrden.AddItem TituloCampo & " [DESC]"
        
        grd.TextMatrix(0, grd.FixedCols + i - 1) = TituloCampo
    Next
End Sub

Public Sub RestaurarEstado()
    Dim Estado() As String
    Dim i As Long
    
    Estado = Split(ReadINI("listas", Format$(m_Lista)), ",")
    
    'Tiene menos de un campo (primera vez o datos mal cargados)
    If UBound(Estado) < 1 Then
        cboCampo.ListIndex = 0
        cboOrden.ListIndex = 0 'Ejecuta el procedimiento 'Mostrar' (por eso no se llama en ningun momento)
'        grd.AutoSize
    Else
        If Val(Estado(0)) < cboCampo.ListCount Then
            cboCampo.ListIndex = Val(Estado(0))
        Else
            cboCampo.ListIndex = 0
        End If
        
        If Val(Estado(1)) < cboOrden.ListCount Then
            cboOrden.ListIndex = Val(Estado(1))
        Else
            cboOrden.ListIndex = 0
        End If
        
        If UBound(Estado) - 2 <= grd.Cols - 1 Then
            For i = 2 To UBound(Estado)
                grd.ColWidth(i - 2) = Val(Estado(i))
            Next
        Else
'            grd.AutoSize
        End If
    End If
End Sub

Private Sub Mostrar()
    Dim rs  As ADODB.Recordset
    Dim sql As String
    
    sql = GetConsulta()
    
    If Len(sql) <> 0 Then
        Set rs = GetRs(sql)
        
        FillGrid grd, rs, True
                    
        If EmptyRS(rs) Then
            cmdMenu(BTN_EDIT).Enabled = False
            cmdMenu(BTN_DELETE).Enabled = False
            picNoData.Visible = True
            sbr.Panels(1).Text = "No hay registros."
        Else
            cmdMenu(BTN_EDIT).Enabled = True
            cmdMenu(BTN_DELETE).Enabled = True
            picNoData.Visible = False
            sbr.Panels(1).Text = str$(rs.RecordCount) & " registros."
        End If
                
        CloseRS rs
    End If
End Sub

Private Function GetConsulta() As String
'    Dim Filtro  As String
'    Dim Orden   As String
'    Dim sql     As String
'
'    sql = m_ObjLista.GetSQL
'
'    Filtro = Trim$(txtFiltro.Text)
'
'    If cboCampo.ListIndex <> -1 And Len(Filtro) <> 0 Then
'        sql = sql & " WHERE " & m_ObjLista.Campos(cboCampo.ListIndex).Nombre & " LIKE " & SQLText("%" & Filter & "%")
'    End If
'
'    If cboOrden.ListIndex <> -1 Then
'        If IsEven(cboOrden.ListIndex) Then
'            sql = sql & " ORDER BY " & StrFormat("{1} ASC", m_ObjLista.Campos(1 + cboOrden.ListIndex / 2).Nombre)
'        Else
'            sql = sql & " ORDER BY " & StrFormat("{1} DESC", m_ObjLista.Campos(1 + (cboOrden.ListIndex - 1) / 2).Nombre)
'        End If
'    End If
    
    GetConsulta = m_ObjLista.GetSQL
End Function

Private Sub IniciarABM(ByVal Accion As EAccionABM)
    Dim f As Form
    Dim Id As String
    Dim Col As Long
    
On Error GoTo ErrorHandler
    If Accion <> abmNuevo And grd.Row = 0 Then
        MsgBox "No hay registros para realizar la acci�n.", vbExclamation
    Else
        Select Case m_Lista
            Case lsProductos:   Set f = New frmProducto
            Case lsRubros:      Set f = New frmRubro
        End Select
        
        If Not Accion = abmNuevo Then
            Id = grd.TextMatrix(grd.Row, grd.FixedCols)
        End If
        
        Select Case Accion
            Case abmNuevo
                f.Iniciar True
                
                If f.ShowModal() = mrOk Then
                    Mostrar
                End If
                
            Case abmModificar
                f.Iniciar False, Id
                
                If f.ShowModal() = mrOk Then
                    Mostrar
                End If
                
            Case abmEliminar
                If f.Eliminar(Id) Then
                    Mostrar
                End If
        End Select
    End If

Finally:
    Set f = Nothing
    
    Exit Sub
ErrorHandler:
    ErrorReport "frmLista", "IniciarABM"
    Resume Finally
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    If Not cmdMenu(Index).Enabled Then
        Exit Sub
    End If

    Select Case Index
        Case BTN_NEW
            IniciarABM abmNuevo

        Case BTN_EDIT
            IniciarABM abmModificar

        Case BTN_DELETE
            IniciarABM abmEliminar

        Case BTN_REFRESH
            Mostrar

        Case BTN_EXPORT
            With New CCommonDialog
                .Init Me
                .Filter = "Libro de Microsoft Office Excel|.xls"
                If .ShowSave() Then
                    If ExportarExcel(.FileName, grd, Me.Caption) Then
                        MsgBox "Datos exportados en '" & .FileName & "'", vbInformation
                    End If
                End If
            End With

        Case BTN_CLOSE
            Unload Me
            
        Case BTN_AUTOSIZE
'            grd.AutoSize
    End Select
End Sub

Private Sub Form_Activate()
On Error Resume Next
    txtFiltro.SetFocus
End Sub

Private Sub Form_Load()
    IniciarLista
    IniciarControles
    RestaurarEstado
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
                     
On Error Resume Next
    grd.Move 60, _
            picFiltro.ScaleHeight + picMenu.ScaleHeight, _
            Me.ScaleWidth - 120, _
            Me.ScaleHeight - picMenu.ScaleHeight - picFiltro.ScaleHeight - sbr.Height - 60
            
    picNoData.Move (Me.Width - picNoData.Width) / 2, _
                   (Me.Height - picNoData.Height) / 2
    
    picOrden.Left = Me.ScaleWidth - picOrden.Width - 60
    txtFiltro.Width = picOrden.Left - txtFiltro.Left - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim Aux As String
    
    Aux = Format$(cboCampo.ListIndex) & "," & Format$(cboOrden.ListIndex) & ","
    
    For i = 0 To grd.Cols - 1
        Aux = Aux & Format$(grd.ColWidth(i)) & ","
    Next
    
    WriteINI "listas", Format$(m_Lista), Left$(Aux, Len(Aux) - 1)
End Sub

Private Sub grd_DblClick()
    If grd.MouseRow <> 0 Then
        cmdMenu_Click BTN_EDIT
    End If
End Sub

Private Sub cboOrden_Click()
    Mostrar
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
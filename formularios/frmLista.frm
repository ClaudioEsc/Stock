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
   Begin ComctlLib.Toolbar tbr 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      _Version        =   327682
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   1635
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2884
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
      MouseIcon       =   "frmLista.frx":058A
   End
   Begin ComctlLib.StatusBar sbr 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
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
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFiltro 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10185
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Width           =   10185
      Begin VB.TextBox txtFiltro 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   300
         Width           =   1815
      End
      Begin VB.ComboBox cboCampo 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1755
      End
      Begin VB.PictureBox picOrden 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   3780
         ScaleHeight     =   555
         ScaleWidth      =   2055
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   2055
         Begin VB.ComboBox cboOrden 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orden"
            ForeColor       =   &H80000010&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda"
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   705
      End
   End
   Begin ComctlLib.ImageList iml 
      Left            =   3060
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":06EC
            Key             =   "actualizar"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":0A3E
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":0D90
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":10E2
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":1434
            Key             =   "exportar"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":1786
            Key             =   "cerrar"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLista.frx":1AD8
            Key             =   "ajustar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Lista     As EListas
Private m_ObjLista  As CLista

Public Sub Iniciar(ByVal Lista As EListas)
    m_Lista = Lista
End Sub

Public Sub IniciarLista()
    Set m_ObjLista = New CLista
    
    Select Case m_Lista
        Case lsProductos
            With m_ObjLista
                .Titulo = "Productos"
                .Tabla = "productos"
                .AgregarCampo "id", "Interno", caDerecha, ctNumero
                .AgregarCampo "codigo", "Código"
                .AgregarCampo "descripcion", "Descripción"
                .AgregarCampo "stock", "Stock", caDerecha, ctNumero, False
                .AgregarCampo "stock_minimo", "Stock mín.", caDerecha, ctNumero, False
            End With
            
        Case lsRubros
            With m_ObjLista
                .Titulo = "Rubros"
                .Tabla = "rubros"
                .AgregarCampo "id", "Interno", caDerecha, ctNumero
                .AgregarCampo "descripcion", "Descripción"
            End With
        
        Case lsMovimientos
            With m_ObjLista
                .Titulo = "Movimientos"
                .Tabla = "movimientos"
                .AgregarCampo "id", "Número", caDerecha, ctNumero
                .AgregarCampo "fecha", "Fecha", caDerecha, ctFecha
                .AgregarCampo "tipo", "Tipo", caCentro, ctTexto
            End With
    End Select
End Sub

Public Sub IniciarControles()
    Dim Campo As CListaCampo
    Dim i As Long
    
    With tbr
        Set .ImageList = iml
        With .Buttons
            .Add , "nuevo", "Nuevo", , "nuevo"
            .Add , "modificar", "Modificar", , "modificar"
            .Add , "eliminar", "Eliminar", , "eliminar"
            .Add , , , tbrSeparator
            
            .Add , "actualizar", "Actualizar", , "actualizar"
            .Add , "ajustar", "Ajustar", , "ajustar"
            .Add , "exportar", "Exportar", , "exportar"
            .Add , , , tbrSeparator
            
            .Add , "cerrar", "Cerrar", , "cerrar"
        End With
    End With
    
    With grd
        .Clear
        .Cols = m_ObjLista.Campos.Count
        .Rows = 1
    End With
    
    i = 0
    
    For Each Campo In m_ObjLista.Campos
        If Campo.PermiteBuscar Then
            With cboCampo
                .AddItem Campo.Titulo
                .ItemData(.NewIndex) = i + 1
            End With
        End If
        
        With cboOrden
            .AddItem Campo.Titulo & " [ASC]"
            .ItemData(.NewIndex) = i + 1
            
            .AddItem Campo.Titulo & " [DESC]"
            .ItemData(.NewIndex) = i + 1
        End With
        
        With grd
            .TextMatrix(0, i) = Campo.Titulo

            Select Case Campo.Alineacion
                Case caDerecha
                    .ColAlignment(i) = flexAlignRightCenter
                    
                Case caCentro
                    .ColAlignment(i) = flexAlignCenterCenter
                    
                Case Else
                    .ColAlignment(i) = flexAlignLeftCenter
            End Select
        End With
                
        i = i + 1
    Next
End Sub

Public Sub RestaurarEstado()
    Dim Estado() As String
    Dim i As Long
    
    Estado = Split(ReadINI("listas", Format$(m_Lista)), ",")
    
    'Tiene menos de un campo (primera vez o datos mal cargados)
    If UBound(Estado) < 1 Then
        cboCampo.ListIndex = 0
        cboOrden.ListIndex = 0 'Ejecuta el procedimiento 'Mostrar'
        GridAutoSize grd
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
            GridAutoSize grd
        End If
    End If
End Sub

Private Sub Mostrar()
    Dim PrevRow     As Long
    Dim StartTime   As Double
    Dim Row         As Long
    Dim Col         As Long
    Dim Campo       As CListaCampo
    Dim rs          As ADODB.Recordset
    Dim Tabla       As String
    
On Error GoTo Catch
    Screen.MousePointer = vbHourglass
    StartTime = Timer()
    
    Tabla = m_ObjLista.Tabla
    Set rs = GetRs(GetConsulta())

    With grd
        .Redraw = False
        PrevRow = .Row
        .Rows = 1
        
        If Not EmptyRS(rs) Then
            .Rows = rs.RecordCount + 1

            For Row = 1 To .Rows - 1
            
                Col = 0
        
                For Each Campo In m_ObjLista.Campos
                    If Len(Campo.Formato) Then
                        .TextMatrix(Row, Col) = Format$(rs.Collect(Campo.Nombre) & vbNullString, Campo.Formato)
                    Else
                        .TextMatrix(Row, Col) = rs.Collect(Campo.Nombre) & vbNullString
                    End If
                    
                    Col = Col + 1
                Next
                
                Select Case Tabla
                    Case "productos"
                        If rs.Collect("stock") < rs.Collect("stock_minimo") Then
                            RowProperty grd, gcpForeColor, &HC0&, Row
                        End If
                End Select
                
                rs.MoveNext
            Next
        End If
        
        .Redraw = True
    End With

    If EmptyRS(rs) Then
        With txtFiltro
            If Len(.Text) Then
                .BackColor = &HC0C0FF
                .ForeColor = vbWhite
            End If
        End With
        
        tbr.Buttons("modificar").Enabled = False
        tbr.Buttons("eliminar").Enabled = False
        sbr.Panels(1).Text = "No se encontraron registros"
    Else
        With txtFiltro
            .BackColor = vbWindowBackground
            .ForeColor = vbButtonText
        End With
        
        tbr.Buttons("modificar").Enabled = True
        tbr.Buttons("eliminar").Enabled = True
        sbr.Panels(1).Text = Format$(rs.RecordCount) & " registros (" & Format$(Timer() - StartTime, "#0.00") & " segundos)"
    End If

Finally:
    CloseRS rs
    GridSelectRow grd, PrevRow
    Screen.MousePointer = vbDefault

    Exit Sub
Catch:
    ErrorReport "frmLista", "Mostrar"
    Resume Finally
End Sub

Private Function GetConsulta() As String
    Dim Filtro  As String
    Dim Orden   As String
    Dim Campo   As CListaCampo
    Dim sql     As String
    
    If cboCampo.ListIndex <> -1 Then
        Filtro = Trim$(txtFiltro.Text)
        
        If Len(Filtro) <> 0 Then
            Set Campo = m_ObjLista.Campos(GetItemData(cboCampo))
            Filtro = Campo.Nombre & " LIKE " & SQLText("%" & Filtro & "%")
        End If
    End If

    If cboOrden.ListIndex <> -1 Then
        Set Campo = m_ObjLista.Campos(GetItemData(cboOrden))
        
        If IsEven(cboOrden.ListIndex) Then
            Orden = Campo.Nombre & " ASC"
        Else
            Orden = Campo.Nombre & " DESC"
        End If
    End If
    
    With New CString
        .Append m_ObjLista.GetSQL
        
        If Len(Filtro) Then
            .Append " WHERE " & Filtro
        End If
        
        If Len(Orden) Then
            .Append " ORDER BY " & Orden
        End If
    
        GetConsulta = .ToString
    End With
End Function

Private Function GetABM() As IFormABM
    Dim f As IFormABM
    
    Select Case m_Lista
        Case lsProductos:   Set f = New frmProducto
        Case lsRubros:      Set f = New frmRubro
        Case lsMovimientos: Set f = New frmMovimiento
    End Select
    
    Set GetABM = f
End Function

Private Function GetId() As String
    With grd
        If .Rows > 1 And .Row > 0 Then
            GetId = grd.TextMatrix(.Row, .FixedCols)
        End If
    End With
End Function

Private Sub NuevoRegistro()
    With GetABM()
        .Iniciar True
                
        If .ShowModal() = mrOK Then
            Mostrar
        End If
    End With
End Sub

Private Sub ModificarRegistro()
    If tbr.Buttons("modificar").Enabled Then
        With GetABM()
            .Iniciar False, GetId()
            
            If .ShowModal() = mrOK Then
                Mostrar
            End If
        End With
    End If
End Sub

Private Sub EliminarRegistro()
    If tbr.Buttons("eliminar").Enabled Then
        If Confirm("¿Desea eliminar el registro seleccionado?", "Eliminar registro", True, False) Then
            With GetABM()
                If .Eliminar(GetId()) Then
                    Mostrar
                End If
            End With
        End If
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    txtFiltro.SetFocus
End Sub

Private Sub Form_Load()
    IniciarLista
    IniciarControles
    RestaurarEstado
    
    Me.Caption = m_ObjLista.Titulo
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
                     
On Error Resume Next
    tbr.Refresh
    
    grd.Move 60, _
            picFiltro.ScaleHeight + tbr.Height, _
            Me.ScaleWidth - 120, _
            Me.ScaleHeight - tbr.Height - picFiltro.ScaleHeight - sbr.Height - 60
                
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
    If grd.MouseRow = 0 Then
        GridAutoSize grd, grd.MouseCol
    Else
        ModificarRegistro
    End If
End Sub

Private Sub cboOrden_Click()
    Mostrar
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:   ModificarRegistro
        Case vbKeyDelete:   EliminarRegistro
        Case vbKeyAdd:      NuevoRegistro
    End Select
End Sub

Private Sub tbr_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "nuevo"
            NuevoRegistro

        Case "modificar"
            ModificarRegistro

        Case "eliminar"
            EliminarRegistro

        Case "actualizar"
            Mostrar

        Case "exportar"
            With New CCommonDialog
                .Init Me
                .Filter = "Libro de Microsoft Office Excel|.xls"
                If .ShowSave() Then
                    If GridExportExcel(grd, .FileName, Me.Caption) Then
                        MsgBox "Datos exportados en '" & .FileName & "'", vbInformation, gAppName
                    End If
                End If
            End With

        Case "cerrar"
            Unload Me
            
        Case "ajustar"
            GridAutoSize grd
    End Select
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

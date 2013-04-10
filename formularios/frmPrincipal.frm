VERSION 5.00
Begin VB.MDIForm frmPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Stock"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8880
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuStock 
      Caption         =   "&Stock"
      Begin VB.Menu mnuProductos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnuRubros 
         Caption         =   "Rubros"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMosaicoHorizontal 
         Caption         =   "Mosaico horizontal"
      End
      Begin VB.Menu mnuMosaicoVertical 
         Caption         =   "Mosaico vertical"
      End
      Begin VB.Menu mnuCascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnuOrganizarIconos 
         Caption         =   "Organizar iconos"
      End
      Begin VB.Menu mnuVentanaLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrarTodas 
         Caption         =   "Cerrar todas"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAcercaDe_Click()
    With New frmAbout
        .Show vbModal, Me
    End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    
    TerminateConnection
End Sub

Private Sub mnuProductos_Click()
    MostrarLista lsProductos
End Sub

Private Sub mnuRubros_Click()
    MostrarLista lsRubros
End Sub

Private Sub MostrarLista(ByVal Lista As EListas)
    With New frmLista
        .Iniciar Lista
        .Show
    End With
End Sub

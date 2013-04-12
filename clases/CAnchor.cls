VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MAX_LAY_BUFF As Long = 20

Private Type TRectangle
    X1  As Long
    Y1  As Long
    X2  As Long
    Y2  As Long
End Type

Private Type TControl
    Control     As Control
    Margin      As TRectangle
    Anchor      As Long
    IsVisible   As Boolean
End Type

Public Enum EAnchorPosition
    apLeft = 1
    apTop = 2
    apRight = 4
    apBottom = 8
    apAll = apLeft + apTop + apRight + apBottom
End Enum

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Private m_Controls()    As TControl
Private m_ControlsSize  As Long
Private m_ControlsCount As Long
Private m_Manual    As Boolean

Public Property Get Manual() As Boolean
    Manual = m_Manual
End Property

Public Property Let Manual(ByVal NewValue As Boolean)
    m_Manual = NewValue
End Property

Private Sub Class_Initialize()
    m_Manual = False
End Sub

Private Sub Class_Terminate()
    Set m_Form = Nothing
    Erase m_Controls
End Sub

Public Sub AddControl(ByRef ctl As Object, ByVal Anchor As EAnchorPosition)
    Dim tmp As TControl
    
    If m_Form Is Nothing Then
        Set m_Form = ctl.Parent
    End If
            
    With tmp
        Set .Control = ctl
        .Margin = GetMargin(ctl)
        .Anchor = Anchor
    End With
    
    If m_ControlsCount = m_ControlsSize Then
        m_ControlsSize = m_ControlsSize + MAX_LAY_BUFF
        ReDim Preserve m_Controls(m_ControlsSize) As TControl
    End If
    
    m_Controls(m_ControlsCount) = tmp
    m_ControlsCount = m_ControlsCount + 1
End Sub

Public Sub Resize()
    Dim i As Long
    For i = 0 To m_ControlsCount - 1
        ResizeControl m_Controls(i)
    Next
End Sub

Private Sub ResizeControl(ByRef lay As TControl)
    Dim Box         As TRectangle
    Dim Position    As TRectangle
    Dim Margin      As TRectangle
    Dim Anchor      As Long
    Dim ctl         As Control
    
    Set ctl = lay.Control

    Box = GetBox(ctl)
    Position = GetPosition(ctl)
    Margin = lay.Margin
    Anchor = lay.Anchor
    
    If apRight = (Anchor And apRight) Then
        If apLeft = (Anchor And apLeft) Then
            Position.X2 = Box.X2 - Margin.X1 - Margin.X2
        Else
            Position.X1 = Box.X2 - Position.X2 - Margin.X2
        End If
    End If
    
    If apBottom = (Anchor And apBottom) Then
        If apTop = (Anchor And apTop) Then
            Position.Y2 = Box.Y2 - Margin.Y1 - Margin.Y2
        Else
            Position.Y1 = Box.Y2 - Position.Y2 - Margin.Y2
        End If
    End If
    
    Call SetPosition(ctl, Position)
End Sub

Private Function GetMargin(ByRef ctl As Control) As TRectangle
    With GetMargin
        .X1 = ctl.Left
        .Y1 = ctl.Top
        .X2 = ctl.Parent.ScaleWidth - .X1 - ctl.Width
        .Y2 = ctl.Parent.ScaleHeight - .Y1 - ctl.Height
    End With
End Function

Private Function GetBox(ByRef ctl As Control) As TRectangle
    With GetBox
        .X1 = ctl.Parent.ScaleLeft
        .Y1 = ctl.Parent.ScaleTop
        .X2 = ctl.Parent.ScaleWidth
        .Y2 = ctl.Parent.ScaleHeight
    End With
End Function

Private Function GetPosition(ByRef ctl As Control) As TRectangle
    With GetPosition
        .X1 = ctl.Left
        .Y1 = ctl.Top
        .X2 = ctl.Width
        .Y2 = ctl.Height
    End With
End Function

Private Function SetPosition(ByRef ctl As Control, ByRef rtg As TRectangle)
    With rtg
        If .X2 < 0 Then .X2 = 0
        If .Y2 < 0 Then .Y2 = 0
        ctl.Left = .X1
        ctl.Top = .Y1
        
        'Algunos controles no permiten modificar estos valores.
        'Ej: La propiedad Height del ComboBox es solo de lectura
        On Error Resume Next
        ctl.Width = .X2
        ctl.Height = .Y2
    End With
End Function

Private Sub m_Form_Resize()
    If Not m_Manual Then
        If m_Form.WindowState <> vbMinimized Then
            Call Resize
        End If
    End If
End Sub

Public Sub Refresh()
    Dim i As Long
    For i = 0 To m_ControlsCount - 1
        With m_Controls(i)
            .Margin = GetMargin(.Control)
        End With
    Next
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************'
'*                                                      *'
'* FileName:    CCommonDialogs.cls                      *'
'* Version:     1.01                                    *'
'* Description: Replacment class for the Common Dialog  *'
'*              Control (ComDlg32.ocx)                  *'
'*              This class also supports the Browse For *'
'*              Folder dialog.                          *'
'* Copyright:   Joacim Andersson 2001                   *'
'*                                                      *'
'*------------------------------------------------------*'
'*                                                      *'
'* Terms of use:                                        *'
'* You may use this class in any application. You may   *'
'* destribute it in compiled or uncompiled form as long *'
'* as these comments aren't altered.                    *'
'*                                                      *'
'********************************************************'

Option Explicit

Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
               PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const LF_FACESIZE = 32
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_USESTYLE = &H80&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const FW_BOLD = 700&
Private Const LOGPIXELSY = 90

Private Declare Function SHBrowseForFolder _
 Lib "shell32" ( _
 lpbi As BrowseInfo) As Long
 
Private Declare Function SHGetPathFromIDList _
 Lib "shell32" ( _
 ByVal pidList As Long, _
 ByVal lpBuffer As String) As Long
 
Private Declare Function lstrcat _
 Lib "kernel32" Alias "lstrcatA" ( _
 ByVal lpString1 As String, _
 ByVal lpString2 As String) As Long
 
Private Declare Function GetOpenFileName _
 Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
 pOpenfilename As OPENFILENAME) As Long
 
Private Declare Function GetSaveFileName _
 Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
 pOpenfilename As OPENFILENAME) As Long
 
Private Declare Function ChooseColor _
 Lib "comdlg32.dll" Alias "ChooseColorA" ( _
 pChoosecolor As ChooseColorStruct) As Long
 
Private Declare Function ChooseFont _
 Lib "comdlg32.dll" Alias "ChooseFontA" ( _
 pChoosefont As ChooseFontStruct) As Long
 
Private Declare Function PrinterProperties _
 Lib "winspool.drv" ( _
 ByVal hWnd As Long, _
 ByVal hPrinter As Long) As Long
 
Private Declare Function OpenPrinter _
 Lib "winspool.drv" Alias "OpenPrinterA" ( _
 ByVal pPrinterName As String, _
 phPrinter As Long, _
 pDefault As PRINTER_DEFAULTS) As Long
 
Private Declare Function ClosePrinter _
 Lib "winspool.drv" ( _
 ByVal hPrinter As Long) As Long
 
Private Declare Function PrintDlg _
 Lib "comdlg32.dll" Alias "PrintDlgA" ( _
 pPrintdlg As PrintDlgStruct) As Long
 
Private Declare Function MulDiv _
 Lib "kernel32" ( _
 ByVal nNumber As Long, _
 ByVal nNumerator As Long, _
 ByVal nDenominator As Long) As Long
 
Private Declare Function GetDeviceCaps _
 Lib "gdi32" ( _
 ByVal hdc As Long, _
 ByVal nIndex As Long) As Long
 
Private Declare Function GetDC _
 Lib "user32" ( _
 ByVal hWnd As Long) As Long
 
Private Declare Function GetDesktopWindow _
 Lib "user32" () As Long
 
Private Declare Function ReleaseDC _
 Lib "user32" ( _
 ByVal hWnd As Long, _
 ByVal hdc As Long) As Long

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type PrintDlgStruct
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As Long
    pDesiredAccess As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Type ChooseFontStruct
    lStructSize As Long
    hwndOwner As Long               '  caller's window handle
    hdc As Long                     '  printer DC/IC or NULL
    lpLogFont As Long
    iPointSize As Long              '  10 * size in points of selected font
    flags As Long                   '  enum. type flags
    rgbColors As Long               '  returned text color
    lCustData As Long               '  data passed to hook fn.
    lpfnHook As Long                '  ptr. to hook function
    lpTemplateName As String        '  custom template name
    hInstance As Long               '  instance handle of.EXE that
                                    '    contains cust. dlg. template
    lpszStyle As String             '  return the style field here
                                    '  must be LF_FACESIZE or bigger
    nFontType As Integer            '  same value reported to the EnumFonts
                                    '    call back with the extra FONTTYPE_
                                    '    bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long                '  minimum pt size allowed &
    nSizeMax As Long                '  max pt size allowed if                                       '    CF_LIMITSIZE is used
End Type

'local variable(s) to hold property value(s)
Private m_FileTitle As String
Private m_Path As String
Private m_Filter As String
Private m_FileName As String
Private m_Color As Long
Private m_Title As String
Private m_Owner As Form
Private m_DefExt As String
Private m_FontName As String
Private m_FontSize As Integer
Private m_FontBold As Boolean
Private m_FontItalic As Boolean
Private m_FontUnderline As Boolean

Public Property Let FontUnderline(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontUnderline = 5
    m_FontUnderline = vData
End Property

Public Property Get FontUnderline() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontUnderline
    FontUnderline = m_FontUnderline
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontItalic = 5
    m_FontItalic = vData
End Property

Public Property Get FontItalic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontItalic
    FontItalic = m_FontItalic
End Property

Public Property Let FontBold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontBold = 5
    m_FontBold = vData
End Property

Public Property Get FontBold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontBold
    FontBold = m_FontBold
End Property

Public Property Let FontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontSize = 5
    m_FontSize = vData
End Property

Public Property Get FontSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontSize
    FontSize = m_FontSize
End Property

Public Property Let FontName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontName = 5
    m_FontName = vData
End Property

Public Property Get FontName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontName
    FontName = m_FontName
End Property

Public Property Let Color(ByVal nNew As Long)
    m_Color = nNew
End Property

Public Property Get Color() As Long
    Color = m_Color
End Property

Public Property Let DefExt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefExt = 5
    m_DefExt = vData
End Property

Public Property Get DefExt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefExt
    DefExt = m_DefExt
End Property

Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    m_Title = vData
End Property

Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Dim nPos As Long
    
    nPos = InStr(m_Title, Chr$(0))
    If nPos Then
        Title = Left$(m_Title, nPos - 1)
    Else
        Title = m_Title
    End If
End Property

Public Sub Init(OwnerForm As Form)
    Set m_Owner = OwnerForm
End Sub

Public Property Let Path(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Path = 5
    m_Path = vData
End Property

Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Dim nPos As Long
    
    nPos = InStr(m_Path, Chr$(0))
    If nPos Then
        Path = Left$(m_Path, nPos - 1)
    Else
        Path = m_Path
    End If
End Property

Public Property Let FileTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileTitle = 5
    m_FileTitle = vData
End Property

Public Property Get FileTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileTitle
    Dim nPos As Long
    
    nPos = InStr(m_FileTitle, Chr$(0))
    If nPos Then
        FileTitle = Left$(m_FileTitle, nPos - 1)
    Else
        FileTitle = m_FileTitle
    End If
End Property

Public Property Get Filter() As String
    Filter = m_Filter
End Property

Public Property Let Filter(ByVal vNewValue As String)
    m_Filter = vNewValue
End Property

Public Property Get FileName() As String
    Dim nPos As Long
    nPos = InStr(m_FileName, Chr$(0))
    If nPos Then
        FileName = Left$(m_FileName, nPos - 1)
    Else
        FileName = m_FileName
    End If
End Property

Public Property Let FileName(sNewVal As String)
    m_FileName = sNewVal
End Property

Public Function ShowPrinter() As Boolean
    Dim pd As PrintDlgStruct
    Dim nRetVal As Long
    
    pd.lpSetupTemplateName = ""
    pd.lpPrintTemplateName = ""
    pd.hwndOwner = m_Owner.hWnd
    pd.flags = PD_HIDEPRINTTOFILE Or PD_NOPAGENUMS Or PD_NOSELECTION
    pd.lStructSize = Len(pd)
    nRetVal = PrintDlg(pd)
    ShowPrinter = (nRetVal <> 0)
End Function

Public Function ShowFolder() As Boolean
    'Opens a Browse Folders Dialog Box that displays the
    'directories in your computer
    Dim lpIDList As Long 'Declare Varibles
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
       .hwndOwner = m_Owner.hWnd 'Owner Form
       .lpszTitle = lstrcat(m_Title, "")
       .ulFlags = BIF_DONTGOBELOWDOMAIN + BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        m_Path = sBuffer
        m_FileName = ""
        m_FileTitle = ""
        ShowFolder = True
    Else
        ShowFolder = False
    End If
End Function

Public Function ShowColor(Optional FullOpen As Boolean = False) As Boolean
    Dim cc As ChooseColorStruct
    Dim nRetVal As Long
    Dim bArr(15) As Byte
    
    cc.lpCustColors = StrConv(bArr, vbUnicode)
    cc.lStructSize = Len(cc)
    cc.flags = IIf(FullOpen, CC_FULLOPEN, 0) Or CC_RGBINIT
    cc.hwndOwner = m_Owner.hWnd
    cc.rgbResult = m_Color
    nRetVal = ChooseColor(cc)
    If nRetVal Then
        m_Color = cc.rgbResult
    End If
    ShowColor = (nRetVal <> 0)
End Function

Public Function ShowFont() As Boolean
    Dim cf As ChooseFontStruct
    Dim nRetVal As Long
    Dim lf As LOGFONT
    Dim nCount As Long
    Dim n As Long
    
    On Error Resume Next
    cf.hwndOwner = m_Owner.hWnd
    cf.flags = CF_BOTH Or CF_INITTOLOGFONTSTRUCT Or CF_FORCEFONTEXIST
    With lf
        If Len(m_FontName) Then
            nCount = Len(m_FontName) - 1
            For n = 0 To nCount
                .lfFaceName(n) = Asc(Mid$(m_FontName, n + 1, 1))
            Next
        End If
        Dim dc As Long
        lf.lfItalic = Abs(m_FontItalic)
        lf.lfUnderline = Abs(m_FontUnderline)
        lf.lfWeight = IIf(m_FontBold, FW_BOLD, 0)
        dc = GetDC(GetDesktopWindow)
        lf.lfHeight = -MulDiv(m_FontSize, GetDeviceCaps(dc, LOGPIXELSY), 72)
        ReleaseDC GetDesktopWindow, dc
    End With
    cf.rgbColors = m_Color
    cf.lpLogFont = VarPtr(lf)
    cf.hdc = Printer.hdc
    cf.lStructSize = Len(cf)
    nRetVal = ChooseFont(cf)
    If nRetVal Then
        m_FontName = StrConv(lf.lfFaceName, vbUnicode)
        m_FontName = Left$(m_FontName, InStr(m_FontName, vbNullChar) - 1)
        m_FontBold = ((cf.nFontType And BOLD_FONTTYPE) = BOLD_FONTTYPE)
        m_FontItalic = ((cf.nFontType And ITALIC_FONTTYPE) = ITALIC_FONTTYPE)
        m_FontSize = cf.iPointSize \ 10
    End If
    ShowFont = (nRetVal <> 0)
End Function

Public Function ShowOpen() As Boolean
    Dim ofn As OPENFILENAME
    Dim nRetVal As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = m_Owner.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = VBA.Replace(m_Filter, "|", Chr$(0)) & Chr$(0)
    ofn.nMaxFile = MAX_PATH + 1
    ofn.nMaxFileTitle = MAX_PATH + 1
    ofn.lpstrFileTitle = Space$(MAX_PATH)
    ofn.lpstrInitialDir = CurDir$
    ofn.lpstrFile = Me.FileName & Space$(MAX_PATH - Len(Me.FileName))
    ofn.lpstrTitle = m_Title
    ofn.flags = &H1000 + &H4 + &H200000 'FileMustExist + HideReadOnly + LongNames
    ofn.lpstrDefExt = m_DefExt
    nRetVal = GetOpenFileName(ofn)
    If nRetVal Then
        m_FileTitle = ofn.lpstrFileTitle
        m_FileName = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
        m_Path = Left$(m_FileName, InStrRev(m_FileName, "\") - 1)
    End If
    ShowOpen = (nRetVal > 0)
End Function

Public Function ShowPrinterProperties() As Boolean
    Dim nRetVal As Long, hPrinter As Long
    Dim pd As PRINTER_DEFAULTS
    
    On Error Resume Next
    pd.pDatatype = 0
    pd.pDesiredAccess = STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE
    pd.pDevMode = 0
    nRetVal = OpenPrinter(Printer.DeviceName, hPrinter, pd)
    If nRetVal <> 0 Then
        nRetVal = PrinterProperties(m_Owner.hWnd, hPrinter)
        ClosePrinter hPrinter
    End If
    ShowPrinterProperties = (nRetVal <> 0)
End Function

Public Function ShowSave() As Boolean
    Dim ofn As OPENFILENAME
    Dim nRetVal As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = m_Owner.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = VBA.Replace(m_Filter, "|", Chr$(0)) & Chr$(0)
    ofn.nMaxFile = MAX_PATH + 1
    ofn.nMaxFileTitle = MAX_PATH + 1
    ofn.lpstrFileTitle = Space$(MAX_PATH)
    ofn.lpstrInitialDir = CurDir$
    ofn.lpstrFile = Me.FileName & Space$(MAX_PATH - Len(Me.FileName))
    ofn.lpstrTitle = m_Title
    ofn.flags = &H800 + &H4 + &H200000 + &H2 'PathMustExist + HideReadOnly + LongNames + OverwritePrompt
    ofn.lpstrDefExt = m_DefExt
    nRetVal = GetSaveFileName(ofn)
    If nRetVal Then
        m_FileTitle = ofn.lpstrFileTitle
        m_FileName = ofn.lpstrFile
        m_Path = Left$(m_FileName, InStrRev(m_FileName, "\") - 1)
    End If
    ShowSave = (nRetVal > 0)
End Function


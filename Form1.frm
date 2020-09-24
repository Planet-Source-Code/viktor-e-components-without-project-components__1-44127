VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CreateWindowEx demonstration"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5715
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbfonts 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3800
      Width           =   1695
   End
   Begin VB.CommandButton bexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create..."
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "UpDown"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "IPAddress"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "TreeView"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Horiz. scroll bar"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ListBox"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ComboBox"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "StatusBar"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ProgressBar"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MonthCalendar"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tabbed dialog"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CommandButton"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "TextBox"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2000
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Label"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2000
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a font:"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label linfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   3420
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CreateWindowEx demonstration
'by Viktor E
'gimelhai@ lycos.com
'With thanks to MSDN :)
'Tested with VB6 EE SP5 on Win98 SE

'You can learn a lot from this example, if you are interested in API. Here are given the names of almost
'all the classes of controls available under Windows, classes which you can manipulate easily to suit
'your needs without explicit reference to the mscomctl or comctl32 libraries (from Project/Components...)
'For example, you might want to create a status bar only 10 minutes, then to make it disappear and
'reappear after another 10 minutes, without redistributing the whole library of controls. For this you
'need to use only 3 generic window style constants, 1 status bar style constant and 3 API functions:
'CreateWindowEx, SendMessage and DestroyWindow;
'the compiled code which uses these gimmicks is considerably smaller than those libraries which you
'might include in your setup package - you can create ImageLists, Toolbars, List-/TreeViews and so on;
'one can create an entire functional interface on-the-fly using only API. You might want to look at
'my other submissions here on PSC, some of them treating the use of control messages subclassing.
'As the MSDN, regarding the action performed by the InitCommonControlsEx function, refers to
'comctl32, which in my Enterprise Edition of VB is referenced as MS Windows Common Controls 5.0,
'fat chance anyway that you'll need to redistribute this library - it is most probably already installed
'by the Windows setup

'A most valuable reference for those who want to learn how to use API is the API Guide and ToolShed
'by Kris and Pieter Philippaerts, or the KPD-Team, tools which might be still available at www.allapi.net
'Do not make fun of the APIs because of .NET, because rapid jumpstarts from Label1.Caption to using
'the .NET seem to be at least unproductive :P

'I preffered to put here all functions, types and constants declaration for easier following of
'the process, as well as for easier adding of new ones by you.
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
    lfFaceName(32) As Byte 'explicit LF_FACESIZE, = 32
End Type
Private Type TVITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type
Private Type tagTVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    item As TVITEM
End Type
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Type tINITCOMMONCONTROLSEX
    'used to initialize common controls library (aka Components/MS Windows Common Controls)
    dwSize As Long
    dwICC As Long
End Type
'We need to initialize common controls library only once; see below why. So, we don't need to declare
'a Boolean variable for use in Form_Load, as it would be useful only for the first program session
Private Type TCITEMHEADER
    'defines a tab control button
    mask As Long
    r1 As Long
    r2 As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef TLPINITCOMMONCONTROLSEX As tINITCOMMONCONTROLSEX) As Long
'^ parameter name modified from INITCOMMONCONTROLSEX to avoid confusion because identical function and type name
'Registers specific common control classes from the common control dynamic-link library
'NOTE: The effect of each call to InitCommonControlsEx is cumulative. For example, if InitCommonControlsEx
'is called with the ICC_UPDOWN_CLASS flag, then is later called with the ICC_HOTKEY_CLASS flag,
'the result is that both the up-down and hot key common control classes are registered and available to the application
'Call to this function is required for MONTHCAL_CLASS
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const ICC_DATE_CLASSES = &H100 'for MONTHCAL
Const ICC_INTERNET_CLASSES = &H800 'for IPADDRESS
'window styles (most of them to use in dwStyle, those specifically for dwStyleEx have an _EX particle):
Const WS_EX_WINDOWEDGE = &H100&
Const WS_EX_STATICEDGE = &H20000
Const WS_EX_CLIENTEDGE = &H200&
Const WS_EX_CONTROLPARENT = &H10000
Const BS_BOTTOM = &H800&
Const BS_LEFT = &H100&
Const BS_MULTILINE = &H2000&
Const ES_MULTILINE = &H4&
Const WS_VSCROLL = &H200000
Const MCS_DAYSTATE = &H1
Const SS_CENTER = &H1&
Const WS_CHILD = &H40000000
Const WS_THICKFRAME = &H40000
Const WS_BORDER = &H800000
Const WS_VISIBLE = &H10000000
Const TCS_FOCUSONBUTTONDOWN = &H1000
Const PBS_SMOOTH = &H1
Const CCS_TOP = &H1& 'StatusBar placed to the top of parent window (dwStyle)
Const CCS_BOTTOM = &H3& 'StatusBar placed to the bottom of parent window (dwStyle)
Const CCS_VERT = &H80& 'StatusBar placed vertical (dwStyle)
Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM) 'StatusBar placed to the right of parent window (dwStyle)
Const CCS_LEFT = (CCS_VERT Or CCS_TOP) 'StatusBar placed to the left of parent window (dwStyle)
Const SBARS_SIZEGRIP = &H100 'OR this in the dwStyle for a sizing grip at the right end of the StatusBar
Const SBT_TOOLTIPS = &H800 'the StatusBar will have tooltips
Const CBS_DROPDOWNLIST = &H3&
Const SBS_HORZ = &H0& 'horizontal scroll bar
Const TVS_HASLINES = &H2
Const TVS_LINESATROOT = &H4
Const TVS_HASBUTTONS = &H1
'^window styles
'common controls class names:
Const ANIMATE_CLASSA = "SysAnimate32"
Const DATETIMEPICK_CLASSA = "SysDateTimePick32"
Const MONTHCAL_CLASSA = "SysMonthCal32"
Const HOTKEY_CLASSA = "msctls_hotkey32"
Const PROGRESS_CLASSA = "msctls_progress32"
Const REBARCLASSNAMEA = "ReBarWindow32"
Const STATUSCLASSNAMEA = "msctls_statusbar32"
Const TOOLBARCLASSNAMEA = "ToolbarWindow32"
Const TOOLTIPS_CLASSA = "tooltips_class32"
Const TRACKBAR_CLASSA = "msctls_trackbar32"
Const UPDOWN_CLASSA = "msctls_updown32"
Const WC_COMBOBOXEXA = "ComboBoxEx32"
Const WC_HEADERA = "SysHeader32"
Const WC_IPADDRESSA = "SysIPAddress32"
Const WC_LISTVIEWA = "SysListView32"
Const WC_PAGESCROLLERA = "SysPager"
Const WC_TABCONTROLA = "SysTabControl32"
Const WC_TREEVIEWA = "SysTreeView32"
'^common controls class names:
'window messages:
Const WM_USER = &H400
Const WM_GETTEXT = &HD
Const WM_SETTEXT = &HC
Const WM_SETFONT = &H30
Const PBM_SETSTEP = (WM_USER + 4)
Const PBM_SETPOS = (WM_USER + 2)
Const TCM_FIRST = &H1300
Const TCM_INSERTITEMA = (TCM_FIRST + 7)
Const TCM_GETCURFOCUS = (TCM_FIRST + 47)
Const TCM_GETCURSEL = (TCM_FIRST + 11)
Const TCM_SETCURFOCUS = (TCM_FIRST + 48)
Const TCM_SETIMAGELIST = (TCM_FIRST + 3)
Const SB_SETPARTS = (WM_USER + 4)
Const SB_SETTEXTA = (WM_USER + 1)
Const SB_SETTIPTEXTA = (WM_USER + 16)
Const LB_ADDSTRING = &H180
Const CB_ADDSTRING = &H143
Const SBM_SETPOS = &HE0
Const SBM_SETRANGE = &HE2 'scrollbar range
Const SBM_SETSCROLLINFO = &HE9
Const SBM_SETRANGEREDRAW = &HE6
Const TV_FIRST = &H1100
Const TVM_INSERTITEMA = (TV_FIRST + 0)
Const TVM_SETINDENT = (TV_FIRST + 7)
Const TVM_SETITEMHEIGHT = (TV_FIRST + 27)
'^window messages
'other constants:
Const TCIF_TEXT = &H1
Const TCIF_IMAGE = &H2
Const TVIF_PARAM = &H4
Const TVIF_TEXT = &H1
Const TVI_FIRST = (-&HFFFF)
Const TVI_ROOT = (-&H10000)
Const SBT_POPOUT = &H200 'raised statusbar panel; effective when OR-ed with SB_SETTEXTA
Const SB_HORZ = 0
'^other constants
'user variables:
Dim tie As TCITEMHEADER, icex As tINITCOMMONCONTROLSEX, si As SCROLLINFO, LogicFont As LOGFONT
Dim h As Long, f As Long
Dim CName As String, WText As String
Dim FSize As Long 'font size
Dim FNameArr() As Byte, NewFont As Long
'^user variables

Private Sub Form_Load()
'To be able to access the MonthCal class we need to register this common control class
'We need to do this only once per program session (ie you may comment this code after the first run);
'see function details^
With icex
    .dwSize = Len(icex)
    .dwICC = ICC_DATE_CLASSES Or ICC_INTERNET_CLASSES
End With
InitCommonControlsEx icex
AddFontNamesToCombo
cbfonts.ListIndex = 1
End Sub
Private Sub bexit_Click()
'exit
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
'clean up the new window
DestroyWindow h
Set Form1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
DestroyWindow h
'^First of all, destroy the previous window, if any.
'Of course that you may complicate yourself checking its existence with "If h<>0 Then"
'That's why no error trapping is set - the window either is created, or not
Select Case Index
Case 0
    CName = "STATIC" 'label
    WText = "This is a Label control" & vbCrLf & vbCrLf & "You can create with code only the entire variety of controls available from the IDE"
    CreateLabel
Case 1
    CName = "EDIT" 'TextBox
    WText = "You can type here. This is a MultiLine TextBox"
    CreateTextBox
Case 2
    CName = "BUTTON" 'CommandButton
    WText = "Yes, I'm a button !" & vbCrLf & "Moreover, a multiline one !" & vbCrLf & "I understand vbCrLf... :P" & vbCrLf & vbCrLf & "Don't stare, click me !"
    CreateCommandButton
Case 3
    CName = "SysTabControl32" 'Tabbed dialog control
    CreateTabControl
Case 4
    CName = "SysMonthCal32" 'Calendar
    CreateMonthCal
Case 5
    CName = "msctls_progress32" 'ProgressBar
    CreateProgressBar
Case 6
    CName = "msctls_statusbar32" 'StatusBar
    CreateStatusBar
Case 7
    CName = "COMBOBOX" 'ComboBox
    CreateComboBox
Case 8
    CName = "LISTBOX" 'ListBox
    CreateListBox
Case 9
    CName = "SCROLLBAR" 'we'll define a horizontal one
    CreateHSBar
Case 10
    CName = "SysTreeView32"
    CreateTreeView
Case 11
    CName = "SysIPAddress32"
    CreateIPAd
Case 12
    CName = "msctls_updown32"
    CreateUpDown
End Select
SetFontFor h, cbfonts.Text
GetInfoAboutTheNewWindow
End Sub

'Now comes the series of control creation procedures. Here's the place where you can add error trapping code,
'mainly because of possible "Bad DLL calling convention" messages when you screw up some parameter data type in SendMessage
Private Sub CreateLabel()
h = CreateWindowEx(WS_EX_WINDOWEDGE, CName, WText, WS_CHILD Or WS_VISIBLE Or SS_CENTER Or WS_THICKFRAME, 170, 20, 160, 120, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
End Sub
Private Sub CreateTextBox()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, WText, WS_CHILD Or WS_BORDER Or WS_VISIBLE Or ES_MULTILINE Or WS_VSCROLL, 170, 20, 160, 120, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
End Sub
Private Sub CreateCommandButton()
h = CreateWindowEx(WS_EX_WINDOWEDGE, CName, WText, BS_MULTILINE Or BS_LEFT Or BS_BOTTOM Or WS_CHILD Or WS_VISIBLE, 170, 20, 160, 120, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
End Sub
Private Sub CreateTabControl()
h = CreateWindowEx(WS_EX_CONTROLPARENT, CName, "", TCS_FOCUSONBUTTONDOWN Or WS_CHILD Or WS_VISIBLE, 170, 20, 200, 120, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
w = Array("Infantry", "Mechanized", "Nukes && biobombs", "Sticks and stones")
For f = 1 To 4
    With tie
        .mask = TCIF_TEXT 'Or TCIF_IMAGE
        .pszText = "WW " & f & "-" & w(f - 1)
        .cchTextMax = 6
        .iImage = -1 'no ImageList set for this td control
    End With
    SendMessage h, TCM_INSERTITEMA, f, tie 'add this button
Next f
End Sub
Private Sub CreateMonthCal()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", WS_BORDER Or WS_CHILD Or WS_VISIBLE Or MCS_DAYSTATE, 170, 20, 200, 160, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
End Sub
Private Sub CreateProgressBar()
'here I surely can define the range using SendMessage h, PBM_SETRANGE, 0, MakeDWord (min,max), but
'for some reason, though, I cannot make it to work properly. Please tell me how it can be done, if you manage to do it
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", PBS_SMOOTH Or WS_CHILD Or WS_VISIBLE, 170, 20, 200, 20, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
SendMessage h, PBM_SETSTEP, 1, 0&
f = 0
Do Until f = 60
    DoEvents
    SendMessage h, PBM_SETPOS, ByVal f, ByVal 0&
    SendMessage h, WM_SETTEXT, 0, ByVal CStr(f)
    f = f + 1
    Sleep 10
Loop
End Sub
Private Sub CreateStatusBar()
Dim SBPartsWidths(3) As Long
SBPartsWidths(0) = 100 'width=100: 100 pixels to the right of the sb left margin
SBPartsWidths(1) = 200 'width=100: the right of this panel is 100 pixels to the right of the right of the previous panel :P
SBPartsWidths(2) = 380 'width=180: the right of this panel is 180 pixels to the right of the right of the previous panel :P
'^for this last panel, one may set its width to the right margin of the parent window by putting -1 instead of a positive value
DestroyWindow h
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", SBT_TOOLTIPS Or WS_CHILD Or WS_VISIBLE, 170, 20, 200, 20, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
SendMessage h, SB_SETPARTS, ByVal 3, SBPartsWidths(0)
For f = 0 To 2
    SendMessage h, SB_SETTEXTA, ByVal f Or SBT_POPOUT, ByVal CStr("Put mouse pointer over this panel") 'SBT_POPOUT is OR-ed with f to raise the text
    If f < 2 Then SendMessage h, SB_SETTIPTEXTA, ByVal f, ByVal CStr("Panel " & f + 1 & " is too short to display all text; otherwise, this tooltip wouldn't appear")
    '^panel 3 won't display tooltips in Tahoma(size=8), because all its text fits inside it
Next f
Erase SBPartsWidths
End Sub
Private Sub CreateComboBox()
h = CreateWindowEx(WS_EX_WINDOWEDGE, CName, "", CBS_DROPDOWNLIST Or WS_VSCROLL Or WS_CHILD Or WS_VISIBLE, 170, 20, 160, 140, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
AddItemsToList h, CB_ADDSTRING
End Sub
Private Sub CreateListBox()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", WS_BORDER Or WS_VSCROLL Or WS_CHILD Or WS_VISIBLE, 170, 20, 160, 120, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
AddItemsToList h, LB_ADDSTRING
End Sub
Private Sub CreateHSBar()
h = CreateWindowEx(WS_EX_WINDOWEDGE, CName, "", SBS_HORZ Or WS_CHILD Or WS_VISIBLE, 170, 20, 200, 20, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
With si
    .cbSize = Len(si)
    .nMax = 10
    .nMin = 0
    .nPos = 3
End With
SetScrollInfo h, SB_HORZ, si, True
SendMessage h, SBM_SETRANGE, ByVal 0, ByVal 10
SendMessage h, SBM_SETRANGEREDRAW, ByVal 0, ByVal 10
For f = 1 To 10
    DoEvents
    SendMessage h, SBM_SETPOS, f, True
    Sleep 100
Next f
'this scroll bar won't respond to clicks...
End Sub
Private Sub CreateTreeView()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", TVS_HASBUTTONS Or TVS_HASLINES Or TVS_LINESATROOT Or WS_BORDER Or WS_CHILD Or WS_VISIBLE, 170, 20, 200, 200, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
SendMessage h, TVM_SETINDENT, 40, 0 'self-explanatory; negative values considered as 0
Dim ti As TVITEM, tis As tagTVINSERTSTRUCT
Dim hPrevNode As Long 'to know below which node to insert new nodes
For f = 0 To 10
    With ti
        .mask = TVIF_TEXT Or TVIF_PARAM
        .pszText = IIf((f < 1), "Parent node", "Child #" & f)
        .cchTextMax = IIf((f < 1), Len("Parent node"), Len("Child #" & f))
        .lParam = f + 1
    End With
    With tis
        .item = ti
        .hInsertAfter = TVI_FIRST
        'if f<5, then each new node will be a child of the previous one;
        'if greater, then put it below the TVI_ROOT level:
        .hParent = IIf((f < 6), hPrevNode, TVI_ROOT)
    End With
    hPrevNode = SendMessage(h, TVM_INSERTITEMA, 0, tis) 'set the new item insertion position
    SendMessage h, TVM_SETITEMHEIGHT, ByVal 16, 0  'set the height of ALL nodes; 16 or something else, play with it...
Next f
End Sub
Private Sub CreateIPAd()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", WS_BORDER Or WS_CHILD Or WS_VISIBLE, 170, 20, 160, 22, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
End Sub
Private Sub CreateUpDown()
h = CreateWindowEx(WS_EX_CLIENTEDGE, CName, "", WS_CHILD Or WS_VISIBLE, 170, 20, 30, 100, Form1.hwnd, vbNull, App.hInstance, ByVal 0&)
'^the width of 30 is arbitrary - you may as well set it to 0
End Sub


'- - - FUNCTIONS - - -


Public Sub SetFontFor(ByVal hwnd As Long, ByVal NewFontName As String)
'Set a font for the new window, as the system puts the default font
FSize = 8 'or whatever
With LogicFont
    .lfHeight = (FSize * -20) / Screen.TwipsPerPixelY
    FNameArr = StrConv(NewFontName & Chr$(0), vbFromUnicode)
    For f = 0 To Len(NewFontName) '6 is the length of "Tahoma", 15 that of "Times New Roman", 5 that of "Arial" and so on
        LogicFont.lfFaceName(f) = FNameArr(f)
    Next f
    'lfFaceName: null-terminated string that specifies the typeface name of the font. The length of this string must
    'not exceed 32 characters, including the null terminator. The EnumFontFamilies function can be used
    'to enumerate the typeface names of all currently available fonts. If lfFaceName is an empty string,
    'GDI uses the first font that matches the other specified attributes
    '.lfWidth = 0 'the average width, in logical units, of characters in the font. If lfWidth is zero, the aspect ratio of the device is matched against the digitization aspect ratio of the available fonts to find the closest match, determined by the absolute value of the difference
    '.lfItalic = 0 '>=1 or TRUE: italic, 0: not italic
    '.lfUnderline = 0 '>=1 or TRUE: underlined, 0: not underlined
    '.lfStrikeOut = 0 '>=1 or TRUE: strikeout, 0: not strikeout
    '.lfEscapement = 0  'angle, in tenths of degrees, between the escapement vector and the x-axis of the device. The escapement vector is parallel to the base line of a row of text; i.e. 60 degrees=600
    '.lfOrientation = 0 'angle, in tenths of degrees, between each character's base line and the x-axis of the device
    '.lfWeight = 400 'Specifies the weight of the font in the range 0 through 1000. For example, 400 is normal and 700 is bold. If this value is zero, a default weight is used
End With
Erase FNameArr
NewFont = CreateFontIndirect(LogicFont)
'Apply the newly created font to the control text:
SendMessage hwnd, WM_SETFONT, NewFont, 0
End Sub
Private Function MakeDWord(wHi As Integer, wLo As Integer) As Long
'aka Visual C++ MAKELONG
If wHi And &H8000& Then
    MakeDWord = (((wHi And &H7FFF&) * 65536) Or (wLo And &HFFFF&)) Or &H80000000
Else
    MakeDWord = (wHi * 65535) + wLo
End If
End Function
Private Sub GetInfoAboutTheNewWindow()
'just displays new window handle
linfo.Caption = "hWnd = " & h
End Sub
Private Sub AddItemsToList(ByVal hwnd As Long, ByVal ListMessage As Long)
'adds 10 items to the defined list and combo
For f = 1 To 10
    SendMessage hwnd, ListMessage, 0, ByVal CStr("Item #" & f) 'ByVal CStr mandatory for readable text
Next f
End Sub
Private Sub AddFontNamesToCombo()
FontNames = Array("Arial", "Tahoma", "Courier New", "MS Sans Serif", "Times New Roman")
For f = 0 To UBound(FontNames)
    SendMessage cbfonts.hwnd, CB_ADDSTRING, 0, ByVal CStr(FontNames(f))
Next f
End Sub
Private Sub cbfonts_Click()
If cbfonts.ListIndex > -1 Then
    SetFontFor h, cbfonts.Text
End If
End Sub

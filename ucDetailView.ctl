VERSION 5.00
Begin VB.UserControl ucDetailView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
End
Attribute VB_Name = "ucDetailView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'\\****************************************************************************
'\\ ucDetailView - Single file clone of a ListView in Report View
'\\ Credits... a lot of people too many to list them all but;
'\\            Ed Wilk          Bug hunter extraordinaire, responsible for most fixes
'\\            Territop         Comments, hints, guides, info, help...
'\\            Paul Caton       THE Subclasser
'\\            Carles P.V.      ucScrollbar
'\\            LaVolpe, Option Explicit & SelfTaught for their uploads regarding
'\\            usercontrols in general
'\\****************************************************************************

'\\****************************************************************************
'\\ Subclassing stuff - Mainly used to trap the mousewheel event **************
    Private Enum eMsgWhen
        MSG_BEFORE = 1
        MSG_AFTER = 2
        MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
    End Enum
    Private Const ALL_MESSAGES  As Long = -1
    Private Const MSG_ENTRIES   As Long = 32
    Private Const WNDPROC_OFF   As Long = &H38
    Private Const GWL_WNDPROC   As Long = -4
    Private Const IDX_SHUTDOWN  As Long = 1
    Private Const IDX_HWND      As Long = 2
    Private Const IDX_WNDPROC   As Long = 9
    Private Const IDX_BTABLE    As Long = 11
    Private Const IDX_ATABLE    As Long = 12
    Private Const IDX_PARM_USER As Long = 13
    Private z_ScMem             As Long
    Private z_Sc(64)            As Long
    Private z_Funk              As Collection
    Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
'    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'\\ End Subclassing stuff *****************************************************
'\\****************************************************************************

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Const WM_TIMER          As Long = &H113
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const C_ROW_SPACING     As Long = 4
Private Const C_ROW_HEIGHT      As Long = 17
Private Const C_TEXT_PADDING    As Long = 5
Private Const C_SCB_SIZE        As Long = 16
Private Const C_SCB_GREY_COLOR  As Long = &HE3E7E9

Private m_lGridLineColor        As Long
Private m_lTextHeight           As Long
Private m_lHeaderBackColor      As Long
Private m_lHeaderDrawStyle      As Long
Private m_lHeaderHeight         As Long
Private m_lHeaderResizeIndex    As Long
Private m_bHeaderResize         As Boolean
Private m_lScrollbarDrawStyle   As Long
Private m_lScrollbarColor       As Long
Private m_sHeaderProperties     As String

Private uLst    As uListInfo
Private uHdr()  As uHeaderInfo
Private uItm()  As uItemInfo
Private uScb(1) As uScrollbarInfo

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type uHeaderInfo
    Caption           As String
    Down              As Boolean
    Rectangle         As RECT
    TextAlign         As AlignmentConstants
    TextBold          As Boolean
    TextColor         As Long
    ListTextAlign     As AlignmentConstants
    ListTextBold      As Boolean
    ListTextColor     As Long
    Sorted            As Boolean
    SortDirectionDown As Boolean
    Width             As Long
End Type

Private Type uHitTestInfo
    ControlID         As Long
    ColumnIndex       As Long
    ListIndex         As Long
End Type

Private Type uListInfo
    BackColor          As Long
    ColumnCount        As Long
    ColumnIndex        As Long
    FullRowSelect      As Boolean
    GridLines          As Boolean
    List()             As Long
    ListCount          As Long
    ListIndex          As Long
    ListPosition       As Long
    Locked             As Boolean
    Rectangle          As RECT
    ListCapacity       As Long
    TopIndex           As Long
    Width              As Long
    HighlightColor     As Long
    HighlightTextColor As Long
End Type

Public Enum eDrawStyle
    Standard = 0
    Soft = 1
    Fill = 2
    Flat = 3
    Raised = 4
    WinXP = 5
    Test1 = 8
End Enum

Private Type uItemInfo
    Items() As String
End Type

Public Enum eHeaderIndex
    Header0 = 0
    Header1 = 1
    Header2 = 2
    Header3 = 3
    Header4 = 4
    Header5 = 5
    Header6 = 6
    Header7 = 7
    Header8 = 8
    Header9 = 9
    Header10 = 10
    Header11 = 11
    Header12 = 12
    Header13 = 13
    Header14 = 14
    Header15 = 15
End Enum

Private Enum eControlID
    scbVUp = 1
    scbVDown = 2
    scbVBlock = 3
    scbVBack = 4
    scbHUp = 5
    scbHDown = 6
    scbHBlock = 7
    scbHBack = 8
    Header = 9
    Body = 10
    Miss = 11
End Enum

Private Type uScrollbarInfo
    DownIndex         As Long
    Drag              As Boolean
    DragOffset        As Long
    Rectangle(3)      As RECT
    TimerStarted      As Boolean
    Visible           As Boolean
End Type

Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Public Event Click(Button As MouseButtonConstants)
Public Event DblClick(Button As MouseButtonConstants)
Public Event HeaderClick(Button As MouseButtonConstants)

'\\****************************************************************************
'\\ Public Properties *********************************************************
Public Property Get BackColor() As OLE_COLOR
    BackColor = uLst.BackColor
End Property
Public Property Let BackColor(xNewData As OLE_COLOR)
    uLst.BackColor = xNewData
    Redraw
    PropertyChanged "lstBackColor"
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = uLst.ColumnCount
End Property
Public Property Let ColumnCount(xNewData As Long)
    uLst.ColumnCount = xNewData
    pvHeaderInit
    pvHeaderRect
    Redraw
    PropertyChanged "lstColumnCount"
End Property

Public Property Get ColumnIndex() As eHeaderIndex
    ColumnIndex = uLst.ColumnIndex
End Property
Public Property Let ColumnIndex(xNewData As eHeaderIndex)
    If xNewData < 0 Then
        uLst.ColumnIndex = 0
    ElseIf xNewData < uLst.ColumnCount Then
        uLst.ColumnIndex = xNewData
    Else
        uLst.ColumnIndex = uLst.ColumnCount - 1
    End If
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = uLst.FullRowSelect
End Property
Public Property Let FullRowSelect(xNewData As Boolean)
    uLst.FullRowSelect = xNewData
    Redraw
    PropertyChanged "lstFullRowSelect"
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = m_lGridLineColor
End Property
Public Property Let GridLineColor(xNewData As OLE_COLOR)
    m_lGridLineColor = xNewData
    Redraw
    PropertyChanged "m_lGridLineColor"
End Property

Public Property Get GridLines() As Boolean
    GridLines = uLst.GridLines
End Property
Public Property Let GridLines(xNewData As Boolean)
    uLst.GridLines = xNewData
    Redraw
    PropertyChanged "lstGridLines"
End Property

Public Property Get ListCount() As Long
    ListCount = uLst.ListCount
End Property

Public Property Let ListIndex(lIndex As Long)
    On Error GoTo errH
    If uLst.ListCount > 0 Then
        If lIndex < uLst.ListCount Then
            If lIndex < 0 Then
                uLst.ListIndex = -1
            Else
                uLst.ListIndex = lIndex
            End If
        Else
            uLst.ListIndex = uLst.ListCount - 1
        End If
        pvListDraw
        pvScrollbarDraw
    Else
        uLst.ListIndex = -1
    End If
    Exit Property
errH:
    Debug.Print
    Debug.Print "ListIndex"
    Debug.Print Err.Description
    Debug.Print
End Property
Public Property Get ListIndex() As Long
    If uLst.ListCount = 0 Or uLst.ListIndex < 0 Then
        ListIndex = -1
    Else
        ListIndex = uLst.List(uLst.ListIndex)
    End If
End Property

Public Property Get Locked() As Boolean
    Locked = uLst.Locked
End Property
Public Property Let Locked(xNewData As Boolean)
    uLst.Locked = xNewData
    If Not xNewData Then Redraw
    PropertyChanged "Locked"
End Property

Public Property Let TopIndex(lIndex As Long)
    On Error GoTo errH
    If uLst.ListCount > 0 Then
        If lIndex > uLst.ListCount - uLst.ListCapacity Then
            TopIndex = uLst.ListCount - uLst.ListCapacity
        ElseIf lIndex < 0 Then
            uLst.TopIndex = 0
        Else
            uLst.TopIndex = lIndex
        End If
        pvListDraw
        pvScrollbarDraw
    Else
        uLst.TopIndex = 0
    End If
    Exit Property
errH:
    Debug.Print
    Debug.Print "TopIndex"
    Debug.Print Err.Description
    Debug.Print
End Property
Public Property Get TopIndex() As Long
    TopIndex = uLst.TopIndex
End Property

Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = uLst.HighlightColor
End Property
Public Property Let HighlightColor(xNewData As OLE_COLOR)
    uLst.HighlightColor = xNewData
    Redraw
    PropertyChanged "lstHighlightColor"
End Property

Public Property Get HighlightTextColor() As OLE_COLOR
    HighlightTextColor = uLst.HighlightTextColor
End Property
Public Property Let HighlightTextColor(xNewData As OLE_COLOR)
    uLst.HighlightTextColor = xNewData
    Redraw
    PropertyChanged "lstHighlightTextColor"
End Property

Public Property Get List(lIndex As Long, Optional lColumnIndex As Long = -1) As String
    If uLst.ColumnCount = 0 Then Exit Property
    If uLst.ListCount = 0 Then Exit Property
    If lIndex < 0 Then Exit Property
    If lIndex >= uLst.ListCount Then Exit Property
    If lColumnIndex >= uLst.ColumnCount Then Exit Property
    If lColumnIndex < 0 Then
        Dim lX As Long
        For lX = 0 To uLst.ColumnCount - 1
            List = List & uItm(lX).Items(lIndex) & vbTab
        Next lX
    Else
        List = uItm(lColumnIndex).Items(lIndex) & vbTab
    End If
    List = Strings.Left$(List, Len(List) - 1)
End Property
Public Property Let List(lIndex As Long, Optional lColumnIndex As Long = -1, sItem As String)
    If uLst.ColumnCount = 0 Then Exit Property
    If uLst.ListCount = 0 Then Exit Property
    If lIndex < 0 Then Exit Property
    If lIndex >= uLst.ListCount Then Exit Property
    If lColumnIndex >= uLst.ColumnCount Then Exit Property
    
    If lColumnIndex < 0 Then
        Dim asData() As String, lX As Long
        asData = Split(sItem, vbTab)
        If UBound(asData) <> uLst.ColumnCount - 1 Then ReDim Preserve asData(uLst.ColumnCount - 1)
        For lX = 0 To uLst.ColumnCount - 1
            uItm(lX).Items(uLst.List(lIndex)) = asData(lX)
        Next lX
        Erase asData
    Else
        uItm(lColumnIndex).Items(uLst.List(lIndex)) = sItem
    End If
    pvListDraw
    pvScrollbarDraw
End Property

Public Property Get Text(Optional lColumnIndex As Long = -1) As String
    If uLst.ColumnCount = 0 Then Exit Property
    If uLst.ListCount = 0 Then Exit Property
    If uLst.ListIndex = -1 Then Exit Property
    Text = List(ListIndex, lColumnIndex)
End Property
Public Property Let Text(Optional lColumnIndex As Long = -1, sItem As String)
    If uLst.ListIndex = -1 Then Exit Property
    List(ListIndex, lColumnIndex) = sItem
End Property

'//*********************************************************************
'// Header properties **************************************************
'//*********************************************************************

Public Property Get ScrollBarColor() As OLE_COLOR
    ScrollBarColor = m_lScrollbarColor
End Property
Public Property Let ScrollBarColor(xNewData As OLE_COLOR)
    m_lScrollbarColor = xNewData
    Redraw
    PropertyChanged "m_lScrollbarColor"
End Property

Public Property Get ScrollbarDrawStyle() As eDrawStyle
    ScrollbarDrawStyle = m_lScrollbarDrawStyle
End Property
Public Property Let ScrollbarDrawStyle(xNewData As eDrawStyle)
    m_lScrollbarDrawStyle = xNewData
    Redraw
    PropertyChanged "m_lScrollbarDrawStyle"
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = m_lHeaderBackColor
End Property
Public Property Let HeaderBackColor(xNewData As OLE_COLOR)
    m_lHeaderBackColor = xNewData
    Redraw
    PropertyChanged "m_lHeaderBackColor"
End Property

Public Property Get HeaderCaption() As String
    HeaderCaption = uHdr(uLst.ColumnIndex).Caption
End Property
Public Property Let HeaderCaption(xNewData As String)
    uHdr(uLst.ColumnIndex).Caption = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get HeaderDrawStyle() As eDrawStyle
    HeaderDrawStyle = m_lHeaderDrawStyle
End Property
Public Property Let HeaderDrawStyle(xNewData As eDrawStyle)
    m_lHeaderDrawStyle = xNewData
    If xNewData = WinXP Then HeaderBackColor = &HE4EDEC
    Redraw
    PropertyChanged "HeaderDrawStyle"
End Property

Public Property Get HeaderHeight() As Long
    HeaderHeight = m_lHeaderHeight
End Property
Public Property Let HeaderHeight(xNewData As Long)
    m_lHeaderHeight = xNewData
    pvHeaderRect
    Redraw
    PropertyChanged "m_lHeaderHeight"
End Property
            
Public Property Get ListTextAlign() As AlignmentConstants
    ListTextAlign = uHdr(uLst.ColumnIndex).ListTextAlign
End Property
Public Property Let ListTextAlign(xNewData As AlignmentConstants)
    uHdr(uLst.ColumnIndex).ListTextAlign = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get ListTextBold() As Boolean
    ListTextBold = uHdr(uLst.ColumnIndex).ListTextBold
End Property
Public Property Let ListTextBold(xNewData As Boolean)
    uHdr(uLst.ColumnIndex).ListTextBold = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get ListTextColor() As OLE_COLOR
    ListTextColor = uHdr(uLst.ColumnIndex).ListTextColor
End Property
Public Property Let ListTextColor(xNewData As OLE_COLOR)
    uHdr(uLst.ColumnIndex).ListTextColor = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get HeaderTextAlign() As AlignmentConstants
    HeaderTextAlign = uHdr(uLst.ColumnIndex).TextAlign
End Property
Public Property Let HeaderTextAlign(xNewData As AlignmentConstants)
    uHdr(uLst.ColumnIndex).TextAlign = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get HeaderTextBold() As Boolean
    HeaderTextBold = uHdr(uLst.ColumnIndex).TextBold
End Property
Public Property Let HeaderTextBold(xNewData As Boolean)
    uHdr(uLst.ColumnIndex).TextBold = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get HeaderTextColor() As OLE_COLOR
    HeaderTextColor = uHdr(uLst.ColumnIndex).TextColor
End Property
Public Property Let HeaderTextColor(xNewData As OLE_COLOR)
    uHdr(uLst.ColumnIndex).TextColor = xNewData
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property

Public Property Get HeaderWidth() As Long
    HeaderWidth = uHdr(uLst.ColumnIndex).Width
End Property
Public Property Let HeaderWidth(xNewData As Long)
    uHdr(uLst.ColumnIndex).Width = xNewData
    pvHeaderRect
    Redraw
    PropertyChanged "m_sHeaderProperties"
End Property
'\\ End Public Properties *****************************************************
'\\****************************************************************************


'\\****************************************************************************
'\\ Public Methods ************************************************************
Public Sub AddItem(sItem As String)
    On Error GoTo errH
    Dim asData() As String, lX As Long
    
    'New item added, tell list it is unsorted again. A nice fix by Ed Wilk
    For lX = 0 To uLst.ColumnCount - 1
        uHdr(lX).Sorted = False
    Next lX

    'I'm still using vbTab to separate column values. It's too difficult to recreate the listview method
    asData = Split(sItem, vbTab)

    'Idiot proofing...
    If UBound(asData) <> uLst.ColumnCount - 1 Then
        ReDim Preserve asData(uLst.ColumnCount - 1)
    End If

    'Redim each column to hold the new items
    For lX = 0 To uLst.ColumnCount - 1
        ReDim Preserve uItm(lX).Items(uLst.ListCount)
        uItm(lX).Items(uLst.ListCount) = asData(lX)
    Next lX

    'Add the new index to our list of indexes
    ReDim Preserve uLst.List(uLst.ListCount)
    uLst.List(uLst.ListCount) = uLst.ListCount

    'Update listcount
    uLst.ListCount = uLst.ListCount + 1

    'Clean up
    Erase asData

    Redraw
    Exit Sub
errH:
    Debug.Print
    Debug.Print "AddItem"
    Debug.Print Err.Description
    Debug.Print
End Sub

Public Sub RemoveItem(lIndex As Long)
    On Error GoTo errH
    
    'Idiot proofing...
    If lIndex > uLst.ListCount - 1 Then Exit Sub
    If lIndex < 0 Then Exit Sub
    
    Dim lX As Long, lY As Long
    
    If uLst.ListCount < 2 Then 'That means: ..., -1, 0 & 1
        'Just empty the list
        pvListInit
    Else
        
        'Get real index from sorted index
        'lY = uLst.List(lIndex)
        lY = uLst.List(uLst.ListIndex) 'Fixed by Ed Wilk
        
        'Remove it from the arrays
        For lX = 0 To uLst.ColumnCount - 1
            pvArrayRemoveString uItm(lX).Items, lY
        Next lX
        
       'Then remove from sorted array index
       'pvArrayRemoveLong uLst.List, lIndex
        pvArrayRemoveLong uLst.List, uLst.ListIndex 'Fixed by Ed Wilk

        'Renumber sorted array index
        For lX = 0 To uLst.ListCount - 2
            If uLst.List(lX) > lY Then uLst.List(lX) = uLst.List(lX) - 1
        Next lX
        
        'Update list info
        uLst.ListCount = uLst.ListCount - 1
        TopIndex = uLst.TopIndex
        ListIndex = uLst.ListIndex - 1
    End If
    
    Redraw
    Exit Sub
errH:
    Debug.Print
    Debug.Print "RemoveItem"
    Debug.Print Err.Description
    Debug.Print
End Sub

Public Sub Clear()
    On Error GoTo errH
    pvListInit
    Redraw
    Exit Sub
errH:
    Debug.Print
    Debug.Print "Clear"
    Debug.Print Err.Description
    Debug.Print
End Sub

Public Sub Redraw()
    On Error GoTo errH
    If uLst.ColumnCount = 0 Then Exit Sub
    If uLst.Locked Then Exit Sub
    pvHeaderRect
    pvListRect
    pvScrollbarRect
    pvScrollbarRect

    pvHeaderDraw
    pvListDraw
    pvScrollbarDraw
    
    Exit Sub
errH:
    Debug.Print
    Debug.Print "Redraw"
    Debug.Print Err.Description
    Debug.Print
End Sub
'\\ End Public Methods ********************************************************
'\\****************************************************************************


'\\****************************************************************************
'\\ Private Methods ***********************************************************
Private Sub pvPropertySave()
    If uLst.ColumnCount = 0 Then Exit Sub
    Dim lX As Long, sTxt As String
    sTxt = ","
    m_sHeaderProperties = ""
    For lX = 0 To uLst.ColumnCount - 1
        With uHdr(lX)
            m_sHeaderProperties = m_sHeaderProperties & .Caption & sTxt & .ListTextAlign & sTxt & .ListTextBold & sTxt & .ListTextColor & sTxt & .TextAlign & sTxt & .TextBold & sTxt & .TextColor & sTxt & .Width & "|"
        End With
    Next lX
    m_sHeaderProperties = Strings.Left$(m_sHeaderProperties, Len(m_sHeaderProperties) - 1)
End Sub

Private Sub pvPropertyLoad()
    Dim lX As Long
    Dim asHdr() As String, asHdrInfo() As String
    asHdr = Split(m_sHeaderProperties, "|")
    For lX = 0 To uLst.ColumnCount - 1
        asHdrInfo = Split(asHdr(lX), ",")
        With uHdr(lX)
            .Caption = asHdrInfo(0)
            .ListTextAlign = asHdrInfo(1)
            .ListTextBold = asHdrInfo(2)
            .ListTextColor = asHdrInfo(3)
            .TextAlign = asHdrInfo(4)
            .TextBold = asHdrInfo(5)
            .TextColor = asHdrInfo(6)
            .Width = asHdrInfo(7)
        End With
    Next lX
    Erase asHdr, asHdrInfo
End Sub

'Private Sub TestAPI()
'    'Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
'    On Error GoTo errH
'    Dim sString1 As String
'    Dim sString2 As String
'    Dim asString(1) As String
'    Dim lOld As Long
'
'    asString(0) = "000000"
'    asString(1) = "111111"
'
'    Text1.Text = Text1.Text & _
'                 "asString(0)        : " & asString(0) & vbCrLf & _
'                 "asString(1)        : " & asString(1) & vbCrLf & _
'                 "LenB(asString(0))  : " & LenB(asString(0)) & vbCrLf & _
'                 "LenB(asString(1))  : " & LenB(asString(1)) & vbCrLf & _
'                 "VarPtr(asString(0)): " & VarPtr(asString(0)) & vbCrLf & _
'                 "VarPtr(asString(1)): " & VarPtr(asString(1)) & vbCrLf & _
'                 "StrPtr(asString(0)): " & StrPtr(asString(0)) & vbCrLf & _
'                 "StrPtr(asString(1)): " & StrPtr(asString(1)) & vbCrLf & _
'                 vbCrLf
'
'    asString(0) = "AAAAAAAA"
'    asString(1) = "BBBBBBBB"
'
'    Text1.Text = Text1.Text & _
'                 "asString(0)        : " & asString(0) & vbCrLf & _
'                 "asString(1)        : " & asString(1) & vbCrLf & _
'                 "LenB(asString(0))  : " & LenB(asString(0)) & vbCrLf & _
'                 "LenB(asString(1))  : " & LenB(asString(1)) & vbCrLf & _
'                 "VarPtr(asString(0)): " & VarPtr(asString(0)) & vbCrLf & _
'                 "VarPtr(asString(1)): " & VarPtr(asString(1)) & vbCrLf & _
'                 "StrPtr(asString(0)): " & StrPtr(asString(0)) & vbCrLf & _
'                 "StrPtr(asString(1)): " & StrPtr(asString(1)) & vbCrLf & _
'                 vbCrLf
'
'    sString1 = "This is string 1."
'    sString2 = "This is string 2."
'
'    Text1.Text = Text1.Text & _
'                 "sString1         : " & sString1 & vbCrLf & _
'                 "sString2         : " & sString2 & vbCrLf & _
'                 vbCrLf
'
'    Text1.Text = Text1.Text & _
'                 "lOld = StrPtr(sString2)" & vbCrLf & _
'                 "RtlMoveMemory VarPtr(sString2), VarPtr(sString1), 4" & vbCrLf & _
'                 "RtlMoveMemory VarPtr(sString1), VarPtr(lOld), 4" & vbCrLf & _
'                 vbCrLf
'
'    lOld = StrPtr(sString2)
'    RtlMoveMemory VarPtr(sString2), VarPtr(sString1), 4
'    RtlMoveMemory VarPtr(sString1), VarPtr(lOld), 4
'
'    Text1.Text = Text1.Text & _
'                 "sString1         : " & sString1 & vbCrLf & _
'                 "sString2         : " & sString2 & vbCrLf & _
'                 vbCrLf
'
'    Exit Sub
'errH:
'    Text1.Text = Text1.Text & _
'                 "Error Description: " & Err.Description & vbCrLf & _
'                 "Error Number     : " & Err.Number & vbCrLf & _
'                 vbCrLf
'End Sub

Private Sub pvArrayRemoveString(ByRef asArray() As String, ByVal lPos As Long)
    Dim lUBound As Long
    Dim lStringAddress As Long
    lUBound = UBound(asArray)
    If Not (lPos = lUBound) Then
        'Save memory address of lPos string
        lStringAddress = StrPtr(asArray(lPos))
        'Move all string pointers after lPos one place up
        RtlMoveMemory VarPtr(asArray(lPos)), VarPtr(asArray(lPos + 1)), (lUBound - lPos) * 4
        'Move pointer to be deleted to the end of the array
        RtlMoveMemory VarPtr(asArray(lUBound)), VarPtr(lStringAddress), 4
    End If
    'Delete last item in string array...
    ReDim Preserve asArray(lUBound - 1)
End Sub

Private Sub pvArrayRemoveLong(ByRef alArray() As Long, ByVal lPos As Long)
    Dim lUBound As Long
    lUBound = UBound(alArray)
    If Not (lPos = lUBound) Then
        RtlMoveMemory VarPtr(alArray(lPos)), VarPtr(alArray(lPos + 1)), (lUBound - lPos) * 4
    End If
    ReDim Preserve alArray(lUBound - 1)
End Sub

Private Sub pvArrayReverseLong(ByRef alArray() As Long)
    Dim lUBound As Long
    Dim lTemp As Long
    Dim lX As Long
    lUBound = UBound(alArray)
    For lX = 0 To lUBound \ 2
        lTemp = alArray(lX)
        alArray(lX) = alArray(lUBound - lX)
        alArray(lUBound - lX) = lTemp
    Next lX
End Sub

Private Sub pvArraySortString(ByRef asSortArray() As String, ByRef alSortedIndex() As Long, Optional ByVal IgnoreCase As Boolean = True)
    Dim sVal1 As String, sVal2 As String
    Dim lX As Long, lRow As Long, lMaxRow As Long, lMinRow As Long
    Dim lSwitch As Long, lLimit As Long, lOffset As Long, lZ As Long
    lMaxRow = UBound(asSortArray)
    lMinRow = LBound(asSortArray)
    ReDim alSortedIndex(lMinRow To lMaxRow)
    For lX = lMinRow To lMaxRow
        alSortedIndex(lX) = lX
    Next
    lOffset = lMaxRow \ 2
    Do While lOffset > 0
        lLimit = lMaxRow - lOffset
        Do
            lSwitch = False
            For lRow = lMinRow To lLimit
                lZ = lZ + 1
                sVal1 = asSortArray(alSortedIndex(lRow))
                sVal2 = asSortArray(alSortedIndex(lRow + lOffset))
                If IgnoreCase Then
                    sVal1 = LCase(sVal1)
                    sVal2 = LCase(sVal2)
                End If
                If sVal1 > sVal2 Then
                    lX = alSortedIndex(lRow)
                    alSortedIndex(lRow) = alSortedIndex(lRow + lOffset)
                    alSortedIndex(lRow + lOffset) = lX
                    lSwitch = lRow
                End If
            Next lRow
            lLimit = lSwitch - lOffset
        Loop While lSwitch
        lOffset = lOffset \ 2
    Loop
End Sub

Private Sub pvArrowDraw(uRect As RECT, lDirection As Long, Optional bDown As Boolean = False)
    Const C_Up As String = "0001000,0011100,0111110,1111111"
    Const C_Down As String = "1111111,0111110,0011100,0001000"
    Const C_Left As String = "0001,0011,0111,1111,0111,0011,0001"
    Const C_Right As String = "1000,1100,1110,1111,1110,1100,1000"
    
    Dim asArrow() As String
    Dim lWidth As Long, lHeight As Long
    Dim lPosX As Long, lPosY As Long
    Dim lX As Long, lY As Long
    Dim uPt As POINTAPI, uArrowRect As RECT
    Dim lPen As Long, lPenOld As Long
    
    pvRectCopy uRect, uArrowRect
    
    Select Case lDirection
        Case 0
            asArrow = Split(C_Up, ",")
        Case 1
            asArrow = Split(C_Down, ",")
        Case 2
            asArrow = Split(C_Left, ",")
        Case 3
            asArrow = Split(C_Right, ",")
    End Select
    
    If bDown Then pvRectTransform uArrowRect, 1, 1, 1, 1
    
    lWidth = Len(asArrow(0))
    lHeight = UBound(asArrow)
    
    lPosX = uArrowRect.Left + (uArrowRect.Right - uArrowRect.Left - lWidth) \ 2
    lPosY = uArrowRect.Top + (uArrowRect.Bottom - uArrowRect.Top - lHeight) \ 2
    
    lPen = CreatePen(0, 1, 0)
    lPenOld = SelectObject(UserControl.hDC, lPen)
    
    For lY = 0 To lHeight
        For lX = 1 To lWidth
            If Mid$(asArrow(lY), lX, 1) = "1" Then
                MoveToEx UserControl.hDC, lPosX + lX, lPosY + lY, uPt
                LineTo UserControl.hDC, lPosX + lX - 1, lPosY + lY - 1
            End If
        Next lX
    Next lY
    
    SelectObject UserControl.hDC, lPenOld
    DeleteObject lPen
    
    Erase asArrow
    
End Sub

Private Sub pvDrawRefresh(uRect As RECT)
    Dim tmpRect As RECT
    With tmpRect
        .Top = uRect.Top - 1
        .Bottom = uRect.Bottom + 1
        .Left = uRect.Left - 1
        .Right = uRect.Right + 1
    End With
    RedrawWindow UserControl.hWnd, tmpRect, ByVal 0, 1
End Sub

Private Sub pvDrawEdge(uRect As RECT, eEdgeStyle As eDrawStyle, Optional lBackColor As Long = &HC8D0D4)
    With uRect
        pvDrawFilledRect .Left, .Top, .Right, .Bottom, lBackColor

        If Not eEdgeStyle = eDrawStyle.Fill Then
            Dim lRed As Long, lGreen As Long, lBlue As Long
            Dim lCA1 As Long, lCA2 As Long, lCA3 As Long, lCB1 As Long, lCB2 As Long

            pvSplitRGB lBackColor, lRed, lGreen, lBlue
            lCA1 = RGB(lRed * 0.35, lGreen * 0.35, lBlue * 0.35) 'Darker
            lCA2 = RGB(lRed * 0.7, lGreen * 0.7, lBlue * 0.7)    'Dark
            lCA3 = RGB(lRed * 1.25, lGreen * 1.25, lBlue * 1.25) 'Lighter

            Select Case eEdgeStyle
                Case eDrawStyle.Standard
                    pvDrawLine .Right - 1, .Top, .Right - 1, .Bottom, lCA2
                    pvDrawLine .Left, .Bottom - 1, .Right, .Bottom - 1, lCA2
                    pvDrawLine .Left, .Top, .Right, .Top, lCA3
                    pvDrawLine .Left, .Top, .Left, .Bottom, lCA3
                    pvDrawLine .Right, .Top, .Right, .Bottom, lCA1
                    pvDrawLine .Left, .Bottom, .Right, .Bottom, lCA1

                Case eDrawStyle.Raised
                    pvDrawLine .Left + 1, .Top + 1, .Left + 1, .Bottom, lCA3
                    pvDrawLine .Left + 1, .Top + 1, .Right, .Top + 1, lCA3
                    pvDrawLine .Right - 1, .Top + 1, .Right - 1, .Bottom, lCA2
                    pvDrawLine .Left + 1, .Bottom - 1, .Right, .Bottom - 1, lCA2
                    pvDrawLine .Right, .Top, .Right, .Bottom, lCA1
                    pvDrawLine .Left, .Bottom, .Right, .Bottom, lCA1
                
                Case eDrawStyle.Soft
                    pvDrawLine .Left, .Top, .Right, .Top, lCA3
                    pvDrawLine .Left, .Top, .Left, .Bottom, lCA3
                    pvDrawLine .Right, .Top, .Right, .Bottom, lCA1
                    pvDrawLine .Left, .Bottom, .Right, .Bottom, lCA1

                Case eDrawStyle.Flat
                    pvDrawLine .Left, .Top, .Right, .Top, lCA2
                    pvDrawLine .Right, .Top, .Right, .Bottom, lCA2
                    pvDrawLine .Left, .Bottom, .Right, .Bottom, lCA2
                    pvDrawLine .Left, .Top, .Left, .Bottom, lCA2

                Case eDrawStyle.WinXP, 7, 6
                    Select Case eEdgeStyle
                        Case eDrawStyle.WinXP
                            lCA1 = RGB(lRed * 1.15, lGreen * 1.15, lBlue * 1.15)
                            lCA2 = RGB(lRed * 1.1, lGreen * 1.1, lBlue * 1.1)
                            lCA3 = RGB(lRed * 1.05, lGreen * 1.05, lBlue * 1.05)
                            lCB1 = RGB(lRed * 0.96, lGreen * 0.96, lBlue * 0.96)
                            lCB2 = RGB(lRed * 0.91, lGreen * 0.91, lBlue * 0.91)
                        Case 7 'WinXP Down
                            lCA1 = RGB(lRed * 0.88, lGreen * 0.88, lBlue * 0.88)
                            lCA2 = RGB(lRed * 0.91, lGreen * 0.91, lBlue * 0.91)
                            lCA3 = RGB(lRed * 0.96, lGreen * 0.96, lBlue * 0.96)
                            lCB1 = RGB(lRed * 1.05, lGreen * 1.05, lBlue * 1.05)
                            lCB2 = RGB(lRed * 1.1, lGreen * 1.1, lBlue * 1.1)
                        Case 6 'WinXP Over?
                            lCA1 = &HC1CCD1
                            lCA2 = &HC8D2D6
                            lCA3 = &HCED8DA
                            lCB1 = &HE1E7E8
                            lCB2 = &HE8ECED
                    End Select
                    pvDrawLine .Left, .Top + 0, .Right, .Top + 0, lCA1        '3 upper gradient h lines
                    pvDrawLine .Left, .Top + 1, .Right, .Top + 1, lCA2        '3 upper gradient h lines
                    pvDrawLine .Left, .Top + 2, .Right, .Top + 2, lCA3        '3 upper gradient h lines
                    pvDrawFilledRect .Left, .Top + 3, .Right, .Bottom - 1, lBackColor 'Single color box
                    pvDrawLine .Left, .Bottom - 1, .Right, .Bottom - 1, lCB1 '2 lower gradient h lines
                    pvDrawLine .Left, .Bottom - 0, .Right, .Bottom - 0, lCB2  '2 lower gradient h lines
                    pvDrawLine .Left, .Top + 2, .Left, .Bottom - 1, lCA1         'left gradient v line
                    pvDrawLine .Right, .Top + 2, .Right, .Bottom - 1, lCB2       'right gradient v line

                Case eDrawStyle.Test1
                    pvDrawLine .Right, .Top, .Right, .Bottom, lCA1
                    pvDrawLine .Right, .Bottom, .Left, .Bottom, lCA1
                    pvDrawLine .Left, .Top, .Right - 1, .Top, lCA3
                    pvDrawLine .Left, .Top, .Left, .Bottom - 1, lCA3
            End Select
        End If
    End With
End Sub

Private Sub pvSplitRGB(ByVal lColor As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)
    lRed = lColor And &HFF
    lGreen = (lColor And &HFF00&) \ &H100&
    lBlue = (lColor And &HFF0000) \ &H10000
End Sub

Private Static Sub pvDrawFilledRect(lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, lColor As Long)
    On Error GoTo errH
    Dim lBrush As Long
    Dim uRect As RECT
    With uRect
        .Left = lLeft
        .Right = lRight + 1
        .Top = lTop
        .Bottom = lBottom + 1
    End With
    lBrush = CreateSolidBrush(lColor)
    FillRect UserControl.hDC, uRect, lBrush
    DeleteObject lBrush
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvDrawFilledRect"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Static Sub pvDrawLine(lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, lColor As Long)
    On Error GoTo errH
    
    Dim uPt As POINTAPI
    Dim lPen As Long, lPenOld As Long
    
    lPen = CreatePen(0, 1, lColor)
    lPenOld = SelectObject(UserControl.hDC, lPen)
    
    If lLeft = lRight And lTop = lBottom Then       'Dot
        MoveToEx UserControl.hDC, lLeft, lTop, uPt
        LineTo UserControl.hDC, lRight + 1, lBottom + 1
        
    ElseIf lLeft = lRight Then                      'Vertical Line
        MoveToEx UserControl.hDC, lLeft, lTop, uPt
        LineTo UserControl.hDC, lRight, lBottom + 1
        
    ElseIf lTop = lBottom Then                      'Horizontal Line
        MoveToEx UserControl.hDC, lLeft, lTop, uPt
        LineTo UserControl.hDC, lRight + 1, lBottom
        
    Else                                            'Diagonal line
        MoveToEx UserControl.hDC, lLeft, lTop, uPt
        LineTo UserControl.hDC, lRight + 1, lBottom + 1
        
    End If
    
    SelectObject UserControl.hDC, lPenOld
    DeleteObject lPen
    
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvDrawLine"
    Debug.Print Err.Description
    Debug.Print
End Sub

'Private Static Sub pvDrawLine(lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, lColor As Long)
'    On Error GoTo errH
'    Dim uPt As POINTAPI
'    Dim lPen As Long
'    Dim lPenOld As Long
'    lPen = CreatePen(0, 1, lColor)
'    lPenOld = SelectObject(UserControl.hDC, lPen)
'    MoveToEx UserControl.hDC, lLeft, lTop, uPt
'    LineTo UserControl.hDC, lRight, lBottom
'    SelectObject UserControl.hDC, lPenOld
'    DeleteObject lPen
'    Exit Sub
'errH:
'    Debug.Print
'    Debug.Print "pvDrawLine"
'    Debug.Print Err.Description
'    Debug.Print
'End Sub

Private Sub pvListRect()
    On Error GoTo errH
    If uLst.ColumnCount = 0 Then Exit Sub

    With uLst.Rectangle
        'Init the body rectangle
        .Top = uHdr(0).Rectangle.Bottom + 1
        .Bottom = UserControl.ScaleHeight
        .Left = 0
        .Right = UserControl.ScaleWidth

        'Fill in additional info
        uLst.ListCapacity = (.Bottom - .Top) \ C_ROW_HEIGHT
    End With
    Exit Sub

errH:
    Debug.Print
    Debug.Print "pvListRect"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvListInit()
    On Error GoTo errH
    Dim lX As Long

    'Clear all arrays
    For lX = 0 To uLst.ColumnCount - 1
        ReDim uItm(lX).Items(0)
    Next lX

    'Clear sorted list index
    ReDim uLst.List(0)

    'Update list info
    uLst.ListCount = 0
    uLst.ListIndex = -1
    uLst.ListPosition = 0
    uLst.TopIndex = 0

    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvListInit"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvListDraw()
    On Error GoTo errH
    Dim lX As Long
    With uLst.Rectangle
        pvDrawFilledRect .Left, .Top, .Right - 1, .Bottom, uLst.BackColor
    End With
    For lX = 0 To uLst.ListCapacity
        pvItemDraw uLst.TopIndex + lX, lX
    Next lX
    pvGridLinesDraw
    pvDrawRefresh uLst.Rectangle
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvListDraw"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvGridLinesDraw()
    On Error GoTo errH
    If uLst.GridLines Then
        Dim lX As Long, lY As Long
        For lX = 1 To uLst.ListCapacity
            lY = uLst.Rectangle.Top + C_ROW_HEIGHT * lX
            pvDrawLine uLst.Rectangle.Left, lY, uLst.Rectangle.Right, lY, m_lGridLineColor
        Next lX
        For lX = 0 To uLst.ColumnCount - 1
            pvDrawLine uHdr(lX).Rectangle.Right, uLst.Rectangle.Top, uHdr(lX).Rectangle.Right, uLst.Rectangle.Bottom, m_lGridLineColor
        Next lX
    End If
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvGridLinesDraw"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Static Sub pvItemDraw(lIndex As Long, lListPosition As Long)
    On Error GoTo errH
        
    If lIndex < 0 Then Exit Sub
    If lIndex > uLst.ListCount - 1 Then Exit Sub
    
    Dim lX As Long, lTextAlignment As Long
    Dim uTR As RECT
    
    'Update text rect
    uTR.Top = uLst.Rectangle.Top + C_ROW_HEIGHT * lListPosition
    uTR.Bottom = uTR.Top + C_ROW_HEIGHT
    
    For lX = 0 To uLst.ColumnCount - 1
    
        With uHdr(lX)
            
            'Get text alignment to use
            If .ListTextAlign = vbCenter Then
                lTextAlignment = 37 'DT_VCENTER Or DT_SINGLELINE Or DT_CENTER
            ElseIf .ListTextAlign = vbLeftJustify Then
                lTextAlignment = 36 'DT_VCENTER Or DT_SINGLELINE Or DT_LEFT
            ElseIf .ListTextAlign = vbRightJustify Then
                lTextAlignment = 38 'DT_VCENTER Or DT_SINGLELINE Or DT_RIGHT
            End If
            
            'Clear the rect
            pvDrawFilledRect .Rectangle.Left, uTR.Top, .Rectangle.Right, uTR.Bottom, uLst.BackColor
            
            'Update text rect
            uTR.Left = .Rectangle.Left + C_TEXT_PADDING
            uTR.Right = .Rectangle.Right - C_TEXT_PADDING
            
            'Set bold
            UserControl.FontBold = .ListTextBold
            
            'If its the selected index
            If uLst.ListIndex = lIndex And (uLst.ColumnIndex = lX Or uLst.FullRowSelect) Then
                pvDrawFilledRect .Rectangle.Left, uTR.Top, .Rectangle.Right, uTR.Bottom, uLst.HighlightColor
                UserControl.ForeColor = uLst.HighlightTextColor
            Else
                pvDrawFilledRect .Rectangle.Left, uTR.Top, .Rectangle.Right, uTR.Bottom, uLst.BackColor
                UserControl.ForeColor = .ListTextColor
            End If
            
            'And then draw item text
            DrawText UserControl.hDC, uItm(lX).Items(uLst.List(lIndex)), -1, uTR, lTextAlignment
            
        End With
    
    Next lX
    
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvItemDraw"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvHeaderDraw()
    Dim lX As Long, uTextRect As RECT, lTextAlignment As Long, uArrowRect As RECT
    For lX = 0 To uLst.ColumnCount - 1
        With uHdr(lX)
            pvRectCopy .Rectangle, uTextRect
            pvRectCopy .Rectangle, uArrowRect
            pvRectTransform uTextRect, C_TEXT_PADDING, 0, -C_TEXT_PADDING, 0
            
            'Set text rect according to alignment & resize/reposition arrow rect.
            If .TextAlign = vbCenter Then
                lTextAlignment = DT_VCENTER Or DT_SINGLELINE Or DT_CENTER '37 DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
                pvRectTransform uArrowRect, .Width - m_lHeaderHeight, 0, 0, 0
            ElseIf .TextAlign = vbLeftJustify Then
                lTextAlignment = DT_VCENTER Or DT_SINGLELINE Or DT_LEFT '32 DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
                pvRectTransform uArrowRect, .Width - m_lHeaderHeight, 0, 0, 0
            ElseIf .TextAlign = vbRightJustify Then
                lTextAlignment = DT_VCENTER Or DT_SINGLELINE Or DT_RIGHT '38 DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE
                pvRectTransform uArrowRect, 0, 0, m_lHeaderHeight - .Width, 0
            End If
            
            'Draw the edges
            If .Down Then
                pvRectTransform uArrowRect, 1, 1, 1, 1
                pvRectTransform uTextRect, 1, 1, 1, 1
                If m_lHeaderDrawStyle = eDrawStyle.WinXP Then
                    pvDrawEdge .Rectangle, 7, m_lHeaderBackColor
                Else
                    pvDrawEdge .Rectangle, eDrawStyle.Flat, m_lHeaderBackColor
                End If
            Else
                pvDrawEdge .Rectangle, m_lHeaderDrawStyle, m_lHeaderBackColor
            End If
            
            'Draw sort direction arrow
            If .Sorted Then
                If .SortDirectionDown Then
                    pvArrowDraw uArrowRect, 1
                Else
                    pvArrowDraw uArrowRect, 0
                End If
            End If
            
            'Text bold and color
            UserControl.FontBold = .TextBold
            UserControl.ForeColor = .TextColor
            
            DrawText UserControl.hDC, .Caption, -1, uTextRect, lTextAlignment
            pvDrawRefresh .Rectangle
        End With
    Next lX
    If uHdr(uLst.ColumnCount - 1).Rectangle.Right < UserControl.ScaleWidth Then
        With uTextRect
            .Top = uHdr(0).Rectangle.Top
            .Bottom = uHdr(0).Rectangle.Bottom
            .Left = uHdr(uLst.ColumnCount - 1).Rectangle.Right + 1
            .Right = UserControl.ScaleWidth + 10 '- 1
        End With
        pvDrawEdge uTextRect, eDrawStyle.Standard, m_lScrollbarColor
        pvDrawRefresh uTextRect
    End If
    
    UserControl.FontBold = False
    UserControl.ForeColor = 0
End Sub

Private Sub pvHeaderInit()
    Dim lX As Long, lY As Long
    lY = uLst.ColumnCount - 1
    ReDim uItm(lY)
    ReDim Preserve uHdr(lY)
    For lX = 0 To uLst.ColumnCount - 1
        With uHdr(lX)
            .Caption = "Header " & lX + 1
            .ListTextAlign = vbCenter
            .TextAlign = vbCenter
            .Width = 200 + (-1 * lX * 50)
        End With
    Next lX
End Sub

Private Sub pvHeaderRect()
    On Error GoTo errH
    If uLst.ColumnCount = 0 Then Exit Sub
    Dim lX As Long
    If uHdr(0).Rectangle.Left < uLst.Rectangle.Right - uLst.Width - 1 Then uHdr(0).Rectangle.Left = uLst.Rectangle.Right - uLst.Width - 1
    If uHdr(0).Rectangle.Left > 0 Then uHdr(0).Rectangle.Left = 0
    For lX = 0 To uLst.ColumnCount - 1
        With uHdr(lX).Rectangle
            .Bottom = m_lHeaderHeight - 1
            If lX > 0 Then .Left = uHdr(lX - 1).Rectangle.Right + 1
            If uHdr(lX).Width < 20 Then uHdr(lX).Width = 30
            .Right = .Left + uHdr(lX).Width - 1
        End With
    Next lX
    uLst.Width = uHdr(uLst.ColumnCount - 1).Rectangle.Right - uHdr(0).Rectangle.Left
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvHeaderRect"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvScrollbarRect()
    Dim lX As Long
    'Do horizontal
    With uScb(1) 'Horizontal
        If uLst.Width > uLst.Rectangle.Right Then  'We need horizontal scrollbar
            .Visible = True
            'Resize list rect to make space for horizontal scrollbar
            uLst.Rectangle.Bottom = UserControl.ScaleHeight - C_SCB_SIZE - 1
            'We have resized list rect so we must recalc list capacity
            uLst.ListCapacity = (uLst.Rectangle.Bottom - uLst.Rectangle.Top) \ C_ROW_HEIGHT
            'Fill in info for scrollbar rects
            For lX = 0 To 3
                .Rectangle(lX).Bottom = UserControl.ScaleHeight - 1
                .Rectangle(lX).Top = .Rectangle(lX).Bottom - C_SCB_SIZE + 1
            Next lX
            '0, Up
            .Rectangle(0).Left = 0
            .Rectangle(0).Right = .Rectangle(0).Left + C_SCB_SIZE - 1
            '1, Down. Check if vertical scrollbar is visible and adjust accordingly
            .Rectangle(1).Right = UserControl.ScaleWidth - 1
            If uScb(0).Visible Then .Rectangle(1).Right = .Rectangle(1).Right - C_SCB_SIZE
            .Rectangle(1).Left = .Rectangle(1).Right - C_SCB_SIZE + 1
            '3, Grey area
            .Rectangle(3).Left = .Rectangle(0).Right + 1
            .Rectangle(3).Right = .Rectangle(1).Left - 1
        Else
            .Visible = False
            'Reset header
            uHdr(0).Rectangle.Left = 0
            pvHeaderRect
            'Resize list rect fill space
            uLst.Rectangle.Bottom = UserControl.ScaleHeight
            'Clear all rects info. Don't want it to interfere with the hit tests
            For lX = 0 To 3
                pvRectReset .Rectangle(lX)
            Next lX
        End If
    End With
    
    'Redo vertical again coz adding horizontal will sometimes make vertical required
    With uScb(0) 'Vertical
        If uLst.ListCount > uLst.ListCapacity Then   'We need vertical scrollbars
            .Visible = True
            'Resize list rect to make space for v scrollbar
            uLst.Rectangle.Right = UserControl.ScaleWidth - C_SCB_SIZE
            'Fill in info for scrollbar rects
            For lX = 0 To 3
                .Rectangle(lX).Right = UserControl.ScaleWidth - 1
                .Rectangle(lX).Left = .Rectangle(lX).Right - C_SCB_SIZE + 1
            Next lX
            '0, Up
            .Rectangle(0).Top = m_lHeaderHeight
            .Rectangle(0).Bottom = .Rectangle(0).Top + C_SCB_SIZE - 1
            '1, Down. Check if horizontal scrollbar is visible and adjust accordingly
            .Rectangle(1).Bottom = UserControl.ScaleHeight - 1
            If uScb(1).Visible Then .Rectangle(1).Bottom = .Rectangle(1).Bottom - C_SCB_SIZE
            .Rectangle(1).Top = .Rectangle(1).Bottom - C_SCB_SIZE + 1
            '3, Grey area
            .Rectangle(3).Top = .Rectangle(0).Bottom + 1
            .Rectangle(3).Bottom = .Rectangle(1).Top - 1
        Else
            .Visible = False
            'Resize list rect to fill space
            uLst.Rectangle.Right = UserControl.ScaleWidth
            'Clear all rects info. Dont want it to interfere with the hit tests
            For lX = 0 To 3
                pvRectReset .Rectangle(lX)
            Next lX
        End If
    End With
End Sub

Private Sub pvScrollbarDraw()
    On Error GoTo errH
    If uLst.ColumnCount = 0 Then Exit Sub
    Dim lX As Long, lY As Long
    Dim lGreyLen As Long, lBlockLen As Long
    Dim lRed As Long, lGreen As Long, lBlue As Long
    Dim uRect As RECT
    pvSplitRGB m_lScrollbarColor, lRed, lGreen, lBlue
    For lY = 0 To 1
        With uScb(lY)
            If .Visible Then
                'Must reposition grab block before drawing
                If lY = 0 Then
                    lGreyLen = .Rectangle(3).Bottom - .Rectangle(3).Top - 5 '- 6 So that there's always a minimum size for the grab block
                    lBlockLen = lGreyLen * uLst.ListCapacity \ uLst.ListCount + 6
                    .Rectangle(2).Top = .Rectangle(3).Top + lGreyLen * uLst.TopIndex \ uLst.ListCount
                    .Rectangle(2).Bottom = .Rectangle(2).Top + lBlockLen
                Else
                    lGreyLen = .Rectangle(3).Right - .Rectangle(3).Left - 5
                    lBlockLen = lGreyLen * (uLst.Rectangle.Right - uLst.Rectangle.Left) \ uLst.Width + 6
                    .Rectangle(2).Left = .Rectangle(3).Left + lGreyLen * -uHdr(0).Rectangle.Left \ uLst.Width
                    .Rectangle(2).Right = .Rectangle(2).Left + lBlockLen
                End If
                'Draw the grey backbar
                pvDrawEdge .Rectangle(3), Fill, RGB(lRed * 1.09, lGreen * 1.09, lBlue * 1.12)
                pvDrawRefresh .Rectangle(3)
                'Draw the up, down and grab block
                For lX = 0 To 2
                    If .DownIndex = lX Then
                        pvDrawEdge .Rectangle(lX), eDrawStyle.Flat, m_lScrollbarColor
                        pvArrowDraw .Rectangle(lX), lX + lY * 2, True
                    Else
                        pvDrawEdge .Rectangle(lX), m_lScrollbarDrawStyle, m_lScrollbarColor
                        If lX < 2 Then pvArrowDraw .Rectangle(lX), lX + lY * 2, False
                    End If
                    pvDrawRefresh .Rectangle(lX)
                Next lX
            End If
        End With
    Next lY
    
    'Draw the small grey square if both scrollbars are visible
    If uScb(0).Visible Then
        If uScb(1).Visible Then
            pvDrawFilledRect uScb(1).Rectangle(1).Right + 1, uScb(0).Rectangle(1).Bottom + 1, UserControl.ScaleWidth, UserControl.ScaleHeight, m_lScrollbarColor
            With uRect
                .Left = uScb(1).Rectangle(1).Right + 1
                .Top = uScb(0).Rectangle(1).Bottom + 1
                .Right = UserControl.ScaleWidth
                .Bottom = UserControl.ScaleHeight
            End With
            pvDrawRefresh uRect
        End If
    End If
    Exit Sub
errH:
    Debug.Print
    Debug.Print "pvScrollbarDraw"
    Debug.Print Err.Description
    Debug.Print
End Sub

Private Sub pvScrollbarEdge()
    Dim lX As Long, lY As Long
    Dim lRed As Long, lGreen As Long, lBlue As Long
    pvSplitRGB m_lScrollbarColor, lRed, lGreen, lBlue
    For lY = 0 To 1
        If uScb(lY).Visible Then
            pvDrawEdge uScb(lY).Rectangle(3), Fill, RGB(lRed * 1.09, lGreen * 1.09, lBlue * 1.12)
            pvDrawRefresh uScb(lY).Rectangle(3)
            For lX = 0 To 2
                If uScb(lY).DownIndex = lX Then
                    pvDrawEdge uScb(lY).Rectangle(lX), eDrawStyle.Flat, m_lScrollbarColor
                Else
                    pvDrawEdge uScb(lY).Rectangle(lX), m_lScrollbarDrawStyle, m_lScrollbarColor
                End If
                pvDrawRefresh uScb(lY).Rectangle(lX)
            Next lX
        End If
    Next lY
End Sub

Private Sub pvTimerStopAll()
    pvTimerKill eControlID.scbHUp
    pvTimerKill eControlID.scbHDown
    pvTimerKill eControlID.scbHBack
    pvTimerKill eControlID.scbVUp
    pvTimerKill eControlID.scbVDown
    pvTimerKill eControlID.scbVBack
End Sub

Private Sub pvTimerSet(ByVal lTimerID As Long, ByVal lTime As Long)
    SetTimer UserControl.hWnd, lTimerID, lTime, 0
End Sub

Private Sub pvTimerKill(ByVal lTimerID As Long)
    KillTimer UserControl.hWnd, lTimerID
End Sub

Private Sub pvRectCopy(uRectSrc As RECT, uRectDest As RECT)
    With uRectDest
        .Top = uRectSrc.Top
        .Bottom = uRectSrc.Bottom
        .Left = uRectSrc.Left
        .Right = uRectSrc.Right
    End With
End Sub

Private Sub pvRectReset(uRect As RECT)
    With uRect
        .Top = 0
        .Bottom = 0
        .Left = 0
        .Right = 0
    End With
End Sub

Private Sub pvRectTransform(uRect As RECT, lLeft As Long, lTop As Long, lRight As Long, lBottom As Long)
    With uRect
        .Top = .Top + lTop
        .Bottom = .Bottom + lBottom
        .Left = .Left + lLeft
        .Right = .Right + lRight
    End With
End Sub

Private Static Function pvHitTest(uPt As POINTAPI, Optional MouseMove As Boolean = False) As uHitTestInfo
    Dim lX As Long, lY As Long

    'No hit
    pvHitTest.ControlID = eControlID.Miss
    
    'Headers
    For lX = 0 To uLst.ColumnCount - 1
        If pvRectHit(uHdr(lX).Rectangle, uPt.X, uPt.Y) Then
            pvHitTest.ControlID = eControlID.Header
            pvHitTest.ColumnIndex = lX
            Exit For
        End If
    Next lX

    'Scrollbars
    If Not pvHitTest.ControlID = eControlID.Miss Then Exit Function
    For lX = 0 To 1
        If uScb(lX).Visible = True Then
            For lY = 1 To 4
                If pvRectHit(uScb(lX).Rectangle(lY - 1), uPt.X, uPt.Y) Then
                    pvHitTest.ControlID = lY + lX * 4
                    Exit For
                End If
            Next lY
        End If
    Next lX

    'ListItems
    If Not pvHitTest.ControlID = eControlID.Miss Then Exit Function
    If uLst.ListCount = 0 Then Exit Function
    Dim lListPosition As Long
    With uLst
        If pvRectHit(.Rectangle, uPt.X, uPt.Y) Then
        
            'Find out which column user clicked
            For lX = 0 To uLst.ColumnCount - 1
                If uPt.X <= uHdr(lX).Rectangle.Right Then
                    pvHitTest.ColumnIndex = lX
                    Exit For
                End If
            Next lX
            
            'Find out which row user clicked
            lY = uLst.Rectangle.Top + C_ROW_HEIGHT
            For lX = 0 To .ListCapacity
                If uPt.Y <= lY Then
                    lListPosition = lX
                    Exit For
                End If
                lY = lY + C_ROW_HEIGHT
            Next lX
            pvHitTest.ListIndex = .TopIndex + lListPosition
            
            pvHitTest.ControlID = eControlID.Body
        End If
    End With
    
End Function

Private Function pvRectHit(udtRect As RECT, lX As Long, lY As Long) As Boolean
    pvRectHit = False
    With udtRect
        If lX < .Left Then
        ElseIf lX > .Right Then
        ElseIf lY < .Top Then
        ElseIf lY > .Bottom Then
        Else
            pvRectHit = True
        End If
    End With
End Function

Private Sub pvOnMouseDown(Button As MouseButtonConstants)
    If uLst.ColumnCount = 0 Then Exit Sub
    Dim uHitResult As uHitTestInfo
    Dim uPt As POINTAPI
    Dim bDoubleClick As Boolean
    Dim lColumnIndex As Long
    Dim lX As Long
    
    If Button = vbMiddleButton Then
        Button = vbLeftButton
        bDoubleClick = True
    End If
    
    GetCursorPos uPt
    ScreenToClient hWnd, uPt
    
    uHitResult = pvHitTest(uPt)

    Select Case uHitResult.ControlID
        Case eControlID.Header
            If Button = vbLeftButton Then
                If UserControl.MousePointer = vbDefault Then
                    Screen.MousePointer = vbHourglass
                    uHdr(uHitResult.ColumnIndex).Down = True
                    pvHeaderDraw
                    DoEvents
                    If uHdr(uHitResult.ColumnIndex).Sorted Then
                        pvArrayReverseLong uLst.List
                        uHdr(uHitResult.ColumnIndex).SortDirectionDown = Not uHdr(uHitResult.ColumnIndex).SortDirectionDown
                    Else
                        pvArraySortString uItm(uHitResult.ColumnIndex).Items, uLst.List
                        uHdr(uHitResult.ColumnIndex).Sorted = True
                        uHdr(uHitResult.ColumnIndex).SortDirectionDown = True
                        For lX = 0 To uLst.ColumnCount - 1
                            If lX <> uHitResult.ColumnIndex Then
                                uHdr(lX).Sorted = False
                                uHdr(lX).SortDirectionDown = False
                            End If
                        Next lX
                    End If
                Else
                    m_bHeaderResize = True
                End If
                pvListDraw
                pvScrollbarDraw
                Screen.MousePointer = vbDefault
            End If
            RaiseEvent HeaderClick(Button)
        
        Case eControlID.Body
            uLst.ColumnIndex = uHitResult.ColumnIndex
            ListIndex = uHitResult.ListIndex
            pvListDraw
            pvScrollbarDraw
            If bDoubleClick Then
                RaiseEvent DblClick(Button)
            Else
                RaiseEvent Click(Button)
            End If
        
        Case eControlID.scbVUp
            If Button = vbLeftButton Then
                uScb(0).DownIndex = 0
                TopIndex = uLst.TopIndex - 1
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(0).TimerStarted Then
                    pvTimerSet eControlID.scbVUp, 300
                    uScb(0).TimerStarted = True
                End If
            End If
        
        Case eControlID.scbVDown
            If Button = vbLeftButton Then
                uScb(0).DownIndex = 1
                TopIndex = uLst.TopIndex + 1
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(0).TimerStarted Then
                    pvTimerSet eControlID.scbVDown, 300
                    uScb(0).TimerStarted = True
                End If
            End If
            
        Case eControlID.scbVBack
            If Button = vbLeftButton Then
                If uPt.Y < uScb(0).Rectangle(2).Top Then
                    TopIndex = uLst.TopIndex - uLst.ListCapacity
                ElseIf uPt.Y > uScb(0).Rectangle(2).Bottom Then
                    TopIndex = uLst.TopIndex + uLst.ListCapacity
                End If
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(0).TimerStarted Then
                    pvTimerSet eControlID.scbVBack, 300
                    uScb(0).TimerStarted = True
                End If
            End If
            
        Case eControlID.scbHUp
            If Button = vbLeftButton Then
                uScb(1).DownIndex = 0
                uHdr(0).Rectangle.Left = uHdr(0).Rectangle.Left + 1
                pvHeaderRect
                pvHeaderDraw
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(1).TimerStarted Then
                    pvTimerSet eControlID.scbHUp, 300
                    uScb(1).TimerStarted = True
                End If
            End If
        
        Case eControlID.scbHDown
            If Button = vbLeftButton Then
                uScb(1).DownIndex = 1
                uHdr(0).Rectangle.Left = uHdr(0).Rectangle.Left - 1
                pvHeaderRect
                pvHeaderDraw
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(1).TimerStarted Then
                    pvTimerSet eControlID.scbHDown, 300
                    uScb(1).TimerStarted = True
                End If
            End If
            
        Case eControlID.scbHBack
            If Button = vbLeftButton Then
                If uPt.Y < uScb(1).Rectangle(2).Top Then
                    TopIndex = uLst.TopIndex - uLst.ListCapacity
                ElseIf uPt.Y > uScb(1).Rectangle(2).Bottom Then
                    TopIndex = uLst.TopIndex + uLst.ListCapacity
                End If
                pvListDraw
                pvScrollbarDraw
                
                If Not uScb(1).TimerStarted Then
                    pvTimerSet eControlID.scbHBack, 300
                    uScb(1).TimerStarted = True
                End If
            End If
            
        Case eControlID.scbHBlock
            If Button = vbLeftButton Then
                uScb(1).Drag = True
                uScb(1).DragOffset = uPt.X - uScb(0).Rectangle(2).Left
            End If
            
        Case eControlID.scbVBlock
            If Button = vbLeftButton Then
                uScb(0).Drag = True
                uScb(0).DragOffset = uPt.Y - uScb(0).Rectangle(2).Top
            End If
            
    End Select
End Sub

Private Sub pvOnMouseMove()
    Dim uPt As POINTAPI
    Dim uHitResult As uHitTestInfo
    Dim lOffset As Long, lGreyLength As Long
    GetCursorPos uPt
    ScreenToClient hWnd, uPt
    
    uHitResult = pvHitTest(uPt, True)
    
    If uHitResult.ControlID = eControlID.Header Then
        'User moved mouse over left side of header. User wants to resize the previous header
        If uPt.X < uHdr(uHitResult.ColumnIndex).Rectangle.Left + 3 Then
            m_lHeaderResizeIndex = uHitResult.ColumnIndex - 1
            UserControl.MousePointer = vbSizeWE
        ElseIf uPt.X > uHdr(uHitResult.ColumnIndex).Rectangle.Right - 3 Then
            m_lHeaderResizeIndex = uHitResult.ColumnIndex
            UserControl.MousePointer = vbSizeWE
        Else
            UserControl.MousePointer = vbDefault
        End If
    Else
        UserControl.MousePointer = vbDefault
    End If
    
    If uScb(0).Drag Then
        lGreyLength = uScb(0).Rectangle(3).Bottom - uScb(0).Rectangle(3).Top - 5
        TopIndex = (uPt.Y - uScb(0).DragOffset - uScb(0).Rectangle(3).Top) * uLst.ListCount \ lGreyLength
    ElseIf m_bHeaderResize And m_lHeaderResizeIndex > -1 Then
        uHdr(m_lHeaderResizeIndex).Width = uPt.X - uHdr(m_lHeaderResizeIndex).Rectangle.Left
        pvHeaderRect
        pvHeaderDraw
        pvListDraw
        pvScrollbarRect
        pvScrollbarDraw
    ElseIf uScb(1).Drag Then
        With uScb(1)
            lGreyLength = .Rectangle(3).Right - .Rectangle(3).Left - 5
            uHdr(0).Rectangle.Left = -(uPt.X - .DragOffset - .Rectangle(3).Left) * uLst.Width \ lGreyLength
        End With
        pvHeaderRect
        pvHeaderDraw
        pvListDraw
        pvScrollbarDraw
    End If

End Sub

Private Sub pvOnMouseUp()
    If uLst.ColumnCount = 0 Then Exit Sub
    pvTimerStopAll
    m_bHeaderResize = False
    Dim lX As Long
    For lX = 0 To uLst.ColumnCount - 1
        uHdr(lX).Down = False
    Next lX
    uScb(0).DownIndex = -1
    uScb(1).DownIndex = -1
    uScb(0).TimerStarted = False
    uScb(1).TimerStarted = False
    uScb(0).Drag = False
    uScb(1).Drag = False
    
    pvHeaderDraw
    pvScrollbarDraw
End Sub

Private Sub pvOnTimer(wParam As Long)
    pvTimerKill wParam
    
    Select Case wParam
        Case eControlID.scbVUp
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbVUp, 50
        
        Case eControlID.scbVDown
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbVDown, 50
        
        Case eControlID.scbHUp
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbHUp, 20
        
        Case eControlID.scbHDown
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbHDown, 20
        
        Case eControlID.scbVBack
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbVBack, 50
        
        Case eControlID.scbHBack
            pvOnMouseDown vbLeftButton
            pvTimerSet eControlID.scbHBack, 50
        
    End Select
End Sub
'\\ End Private Methods *******************************************************
'\\****************************************************************************

'\\****************************************************************************
'\\ Subclassing stuff *********************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, Optional ByVal lParamUser As Long = 0, Optional ByVal nOrdinal As Long = 1, Optional ByVal oCallback As Object = Nothing, Optional ByVal bIdeSafety As Boolean = True) As Boolean
    Const CODE_LEN      As Long = 260
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))
    Const PAGE_RWX      As Long = &H40&
    Const MEM_COMMIT    As Long = &H1000&
    Const MEM_RELEASE   As Long = &H8000&
    Const IDX_EBMODE    As Long = 3
    Const IDX_CWP       As Long = 4
    Const IDX_SWL       As Long = 5
    Const IDX_FREE      As Long = 6
    Const IDX_BADPTR    As Long = 7
    Const IDX_OWNER     As Long = 8
    Const IDX_CALLBACK  As Long = 10
    Const IDX_EBX       As Long = 16
    Const SUB_NAME      As String = "sc_Subclass"
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    If IsWindow(lng_hWnd) = 0 Then
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    nMyID = GetCurrentProcessId
    GetWindowThreadProcessId lng_hWnd, nID
    If nID <> nMyID Then
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    If oCallback Is Nothing Then
        Set oCallback = Me
    End If
    nAddr = zAddressOf(oCallback, nOrdinal)
    If nAddr = 0 Then
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
    If z_Funk Is Nothing Then
        Set z_Funk = New Collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")
        z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
    If z_ScMem <> 0 Then
        On Error GoTo CatchDoubleSub
        z_Funk.Add z_ScMem, "h" & lng_hWnd
        On Error GoTo 0
        If bIdeSafety Then
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")
        End If
        z_Sc(IDX_EBX) = z_ScMem
        z_Sc(IDX_HWND) = lng_hWnd
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)
        z_Sc(IDX_CALLBACK) = nAddr
        z_Sc(IDX_PARM_USER) = lParamUser
        nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)
        If nAddr = 0 Then
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        z_Sc(IDX_WNDPROC) = nAddr
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN
        sc_Subclass = True
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If
    Exit Function
CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE
End Function
Private Sub sc_Terminate()
    Dim i As Long
    If Not (z_Funk Is Nothing) Then
        With z_Funk
            For i = .Count To 1 Step -1
                z_ScMem = .Item(i)
                If IsBadCodePtr(z_ScMem) = 0 Then
                    sc_UnSubclass zData(IDX_HWND)
                End If
            Next i
        End With
        Set z_Funk = Nothing
    End If
End Sub
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
    If z_Funk Is Nothing Then
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
            zData(IDX_SHUTDOWN) = -1
            zDelMsg ALL_MESSAGES, IDX_BTABLE
            zDelMsg ALL_MESSAGES, IDX_ATABLE
        End If
        z_Funk.Remove "h" & lng_hWnd
    End If
End Sub
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then
            zAddMsg uMsg, IDX_BTABLE
        End If
        If When And MSG_AFTER Then
            zAddMsg uMsg, IDX_ATABLE
        End If
    End If
End Sub
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then
            zDelMsg uMsg, IDX_BTABLE
        End If
        If When And MSG_AFTER Then
            zDelMsg uMsg, IDX_ATABLE
        End If
    End If
End Sub
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
    End If
End Function
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_lParamUser = zData(IDX_PARM_USER)
    End If
End Property
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        zData(IDX_PARM_USER) = NewValue
    End If
End Property
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long
    Dim nBase  As Long
    Dim i      As Long
    nBase = z_ScMem
    z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        nCount = ALL_MESSAGES
    Else
        nCount = zData(0)
        If nCount >= MSG_ENTRIES Then
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If
        For i = 1 To nCount
            If zData(i) = 0 Then
                zData(i) = uMsg
                GoTo Bail
            ElseIf zData(i) = uMsg Then
                GoTo Bail
            End If
        Next i
        nCount = i
        zData(nCount) = uMsg
    End If
    zData(0) = nCount
Bail:
    z_ScMem = nBase
End Sub
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long
    Dim nBase  As Long
    Dim i      As Long
    nBase = z_ScMem
    z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        zData(0) = 0
    Else
        nCount = zData(0)
        For i = 1 To nCount
            If zData(i) = uMsg Then
                zData(i) = 0
                GoTo Bail
            End If
        Next i
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
Bail:
    z_ScMem = nBase
End Sub
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zFnAddr
End Function
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
    If z_Funk Is Nothing Then
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
        On Error GoTo Catch
        z_ScMem = z_Funk("h" & lng_hWnd)
        zMap_hWnd = z_ScMem
    End If
    Exit Function
Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"
End Function
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub  As Byte
    Dim bVal  As Byte
    Dim nAddr As Long
    Dim i     As Long
    Dim j     As Long
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4
    If Not zProbe(nAddr + &H1C, i, bSub) Then
        If Not zProbe(nAddr + &H6F8, i, bSub) Then
            If Not zProbe(nAddr + &H7A4, i, bSub) Then
                Exit Function
            End If
        End If
    End If
    i = i + 4
    j = i + 1024
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4
        If IsBadCodePtr(nAddr) Then
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4
            Exit Do
        End If
        RtlMoveMemory VarPtr(bVal), nAddr, 1
        If bVal <> bSub Then
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4
            Exit Do
        End If
        i = i + 4
    Loop
End Function
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
    nAddr = nStart
    nLimit = nAddr + 32
    Do While nAddr < nLimit
        RtlMoveMemory VarPtr(nEntry), nAddr, 4
        If nEntry <> 0 Then
            RtlMoveMemory VarPtr(bVal), nEntry, 1
            If bVal = &H33 Or bVal = &HE9 Then
                nMethod = nAddr
                bSub = bVal
                zProbe = True
                Exit Function
            End If
        End If
        nAddr = nAddr + 4
    Loop
End Function
Private Property Get zData(ByVal nIndex As Long) As Long
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property
Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property
'\\ End Subclassing stuff *****************************************************
'\\****************************************************************************

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If uLst.ListCount = 0 Then Exit Sub
    
    Select Case KeyCode
    
        Case vbKeyUp
            ListIndex = uLst.ListIndex - 1
            If uLst.TopIndex - ListIndex > 0 Then TopIndex = uLst.TopIndex - 1
            
        Case vbKeyDown
            ListIndex = uLst.ListIndex + 1
            If uLst.ListIndex - uLst.TopIndex > uLst.ListCapacity - 1 Then TopIndex = uLst.TopIndex + 1
            
        Case vbKeyLeft
            If uLst.FullRowSelect Then
                UserControl_KeyDown vbKeyUp, 0
            Else
                uLst.ColumnIndex = uLst.ColumnIndex - 1
                If uLst.ColumnIndex < 0 Then
                    uLst.ColumnIndex = uLst.ColumnCount - 1
                    UserControl_KeyDown vbKeyUp, 0
                Else
                    pvListDraw
                End If
            End If

        Case vbKeyRight
            If uLst.FullRowSelect Then
                UserControl_KeyDown vbKeyDown, 0
            Else
                uLst.ColumnIndex = uLst.ColumnIndex + 1
                If uLst.ColumnIndex > uLst.ColumnCount - 1 Then
                    uLst.ColumnIndex = 0
                    UserControl_KeyDown vbKeyDown, 0
                Else
                    pvListDraw
                End If
            End If
            
        Case vbKeyPageUp
            ListIndex = uLst.ListIndex - uLst.ListCapacity
            TopIndex = uLst.TopIndex - uLst.ListCapacity
            
        Case vbKeyPageDown
            ListIndex = uLst.ListIndex + uLst.ListCapacity
            TopIndex = uLst.TopIndex + uLst.ListCapacity
            
        Case vbKeyEnd
            ListIndex = uLst.ListCount - 1
            TopIndex = uLst.ListCount - uLst.ListCapacity
            
        Case vbKeyHome
            ListIndex = 0
            TopIndex = 0
    
    End Select
    pvListDraw
    pvScrollbarDraw
'    RaiseEvent ItemClick(vbLeftButton, ListIndex, uList.ColumnIndex)
End Sub

'\\****************************************************************************
'\\ UserControl Methods *******************************************************
'\\****************************************************************************

Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub UserControl_Terminate()
    Erase uScb, uHdr, uItm
    sc_Terminate
End Sub

Private Sub UserControl_InitProperties()
    Dim lX As Long
    m_lGridLineColor = &HC8D0D4
    With uLst
        .BackColor = &HFFFFFF
        .ColumnCount = 3
        .FullRowSelect = True
        .GridLines = True
        .HighlightColor = &H800000
        .HighlightTextColor = &HFFFFFF
        .ListIndex = -1
    End With
    uScb(0).DownIndex = -1
    uScb(1).DownIndex = -1
    m_lScrollbarDrawStyle = eDrawStyle.Raised
    m_lScrollbarColor = &HC8D0D4
    m_lHeaderHeight = 20
    m_lHeaderBackColor = &HC8D0D4
    m_lHeaderDrawStyle = eDrawStyle.Standard
    pvHeaderInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    pvPropertySave
    With PropBag
        .WriteProperty "lstBackColor", uLst.BackColor
        .WriteProperty "lstColumnCount", uLst.ColumnCount
        .WriteProperty "lstFullRowSelect", uLst.FullRowSelect
        .WriteProperty "lstGridLines", uLst.GridLines
        .WriteProperty "lstHighlightColor", uLst.HighlightColor
        .WriteProperty "lstHighlightTextColor", uLst.HighlightTextColor
        .WriteProperty "m_sHeaderProperties", m_sHeaderProperties
        .WriteProperty "m_lHeaderHeight", m_lHeaderHeight
        .WriteProperty "m_lHeaderBackColor", m_lHeaderBackColor
        .WriteProperty "m_lHeaderDrawStyle", m_lHeaderDrawStyle
        .WriteProperty "m_lScrollbarColor", m_lScrollbarColor
        .WriteProperty "m_lScrollbarDrawStyle", m_lScrollbarDrawStyle
        .WriteProperty "m_lGridLineColor", m_lGridLineColor
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        uLst.BackColor = .ReadProperty("lstBackColor")
        uLst.ColumnCount = .ReadProperty("lstColumnCount")
        uLst.FullRowSelect = .ReadProperty("lstFullRowSelect")
        uLst.GridLines = .ReadProperty("lstGridLines")
        uLst.HighlightColor = .ReadProperty("lstHighlightColor")
        uLst.HighlightTextColor = .ReadProperty("lstHighlightTextColor")
        m_sHeaderProperties = .ReadProperty("m_sHeaderProperties")
        m_lHeaderHeight = .ReadProperty("m_lHeaderHeight")
        m_lHeaderBackColor = .ReadProperty("m_lHeaderBackColor")
        m_lHeaderDrawStyle = .ReadProperty("m_lHeaderDrawStyle")
        m_lScrollbarColor = .ReadProperty("m_lScrollbarColor")
        m_lScrollbarDrawStyle = .ReadProperty("m_lScrollbarDrawStyle")
        m_lGridLineColor = .ReadProperty("m_lGridLineColor")
    End With
    pvHeaderInit
    pvListInit
    pvPropertyLoad
    uScb(0).DownIndex = -1
    uScb(1).DownIndex = -1
    Redraw
    If Ambient.UserMode Then
        sc_Subclass UserControl.hWnd
        sc_AddMsg UserControl.hWnd, WM_TIMER, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_LBUTTONDOWN, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_LBUTTONUP, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_LBUTTONDBLCLK, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_RBUTTONDOWN, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_MOUSEWHEEL, MSG_AFTER
        sc_AddMsg UserControl.hWnd, WM_MOUSEMOVE, MSG_AFTER
    End If
End Sub

'\\****************************************************************************
'\\ Subclassing stuff *********************************************************
'\\****************************************************************************
Private Sub zWndProc1(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)

    Select Case uMsg
        Case WM_TIMER
            pvOnTimer wParam
        Case WM_LBUTTONDOWN
            pvOnMouseDown vbLeftButton
        Case WM_LBUTTONUP
            pvOnMouseUp
        Case WM_LBUTTONDBLCLK
            pvOnMouseDown vbMiddleButton
        Case WM_RBUTTONDOWN
            pvOnMouseDown vbRightButton
        Case WM_MOUSEWHEEL
            If wParam < 0 Then
                TopIndex = uLst.TopIndex + 1
            ElseIf wParam > 0 Then
                TopIndex = uLst.TopIndex - 1
            End If
        Case WM_MOUSEMOVE
            pvOnMouseMove
    End Select

End Sub

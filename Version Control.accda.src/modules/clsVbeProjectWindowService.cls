VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbeProjectWindowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Const TV_FIRST = &H1100
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_GETITEM = (TV_FIRST + 12)
Private Const TVGN_PARENT = &H3
Private Const TVGN_CARET = &H9

Private Const TREEVIEW_CLASS As String = "SysTreeView32"

Private Type TVITEMEX
    Mask As Long
    hItem As LongPtr
    State As Long
    StateMask As Long
    pszText As LongPtr
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As LongPtr
End Type

Function GetSelectedProjectExplorerItem(ByRef ItemLevel As Long) As String

    Dim hVBE As LongPtr
    Dim hProjectWindow As LongPtr
    Dim ItemText As String

    hVBE = VBE.MainWindow.hwnd

    hProjectWindow = FindWindowEx(hVBE, 0, vbNullString, GetProjectWindow.Caption)
    If hProjectWindow = 0 Then Exit Function

    GetSelectedProjectExplorerItem = GetSelectTreeviewItem(hProjectWindow, ItemLevel)

End Function

Private Function GetProjectWindowHandle(ByVal VbeHandle As LongPtr) As LongPtr

    Dim ProjectWindow As Window
    Set ProjectWindow = GetProjectWindow

    If ProjectWindow Is Nothing Then
        GetProjectWindowHandle = 0
        Exit Function
    End If

    GetProjectWindowHandle = FindWindowEx(VbeHandle, 0, vbNullString, ProjectWindow.Caption)

End Function

Private Function GetProjectWindow() As VBIDE.Window

    Dim w As VBIDE.Window

    For Each w In Application.VBE.Windows
        If w.Type = vbext_wt_ProjectWindow Then
            Set GetProjectWindow = w
            Exit Function
        End If
    Next
    Set GetProjectWindow = Nothing

End Function

Private Function GetSelectTreeviewItem(ByVal hProjectWindowHandle As LongPtr, ByRef ItemLevel As Long) As String

    Dim hTreeView As LongPtr
    Dim hSelectedItem As LongPtr

    hTreeView = FindWindowEx(hProjectWindowHandle, 0, TREEVIEW_CLASS, vbNullString)
    If hTreeView = 0 Then Exit Function

    hSelectedItem = SendMessage(hTreeView, TVM_GETNEXTITEM, TVGN_CARET, 0)
    If hSelectedItem = 0 Then Exit Function

    ItemLevel = GetSelectedTreeviewItemLevel(hTreeView, hSelectedItem)
    GetSelectTreeviewItem = GetSelectedTreeviewItemHandleText(hTreeView, hSelectedItem)

End Function

Private Function GetSelectedTreeviewItemLevel(ByVal hTreeView As LongPtr, ByVal hSelectedItem As LongPtr) As Long

    Dim hParentItem As LongPtr
    Dim Level As Long

    Level = 1

    hParentItem = SendMessage(hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hSelectedItem)

    Do While hParentItem <> 0
        Level = Level + 1
        hParentItem = SendMessage(hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, hParentItem)
    Loop

    GetSelectedTreeviewItemLevel = Level

End Function

Private Function GetSelectedTreeviewItemHandleText(ByVal hTreeView As LongPtr, ByVal hSelectedItem As LongPtr) As String

    Dim tvItem As TVITEMEX
    Dim arrBuffer() As Byte
    Const bufferSize As Long = 256

    ReDim arrBuffer(bufferSize - 1)

    tvItem.hItem = hSelectedItem
    tvItem.Mask = &H1  ' TVIF_TEXT
    tvItem.pszText = VarPtr(arrBuffer(0))
    tvItem.cchTextMax = bufferSize

    Call SendMessage(hTreeView, TVM_GETITEM, 0, VarPtr(tvItem))

    GetSelectedTreeviewItemHandleText = ConvertByteArrayToString(arrBuffer)

End Function

Private Function ConvertByteArrayToString(arrToConvert() As Byte) As String

    Dim i As Long
    Dim strResult As String

    strResult = Space(UBound(arrToConvert) + 1)

    For i = LBound(arrToConvert) To UBound(arrToConvert)
        If arrToConvert(i) = 0 Then Exit For
        Mid(strResult, i + 1, 1) = Chr(arrToConvert(i))
    Next i

    ConvertByteArrayToString = Trim(strResult)

End Function

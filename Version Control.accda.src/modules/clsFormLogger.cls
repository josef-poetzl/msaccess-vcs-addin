VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Console As clsConcat  ' Console output
Private m_RichText As TextBox   ' Text box to display HTML
Private m_Prog As clsLblProg    ' Progress bar
Private m_blnProgressActive As Boolean
Private m_sngLastUpdate As Single
Private m_PadLength As Integer

Private WithEvents m_LogEvents As clsLog
Attribute m_LogEvents.VB_VarHelpID = -1

Implements ILogger

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the log buffers and release object references
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Class_Initialize
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Spacer
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Add a spacer to the log
'---------------------------------------------------------------------------------------
'
Public Sub Spacer()
    Me.Add cstrSpacer
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add a log file entry.
'---------------------------------------------------------------------------------------
'
Public Sub Add(strText As String, _
    Optional blnNextOutputOnNewLine As Boolean = True, _
    Optional strColor As String = vbNullString, _
    Optional blnBold As Boolean = False, _
    Optional blnItalic As Boolean = False)

    ' Remove existing progress indicator if in use.
    If m_blnProgressActive Then
        m_blnProgressActive = False
        m_Prog.Hide
    End If

    ' Build HTML output for console
    With m_Console

        ' Opening tags
        If blnBold Then .Add "<strong>"
        If blnItalic Then .Add "<i>"
        If strColor <> vbNullString Then .Add "<font color=", strColor, ">"

        ' Content
        .Add MultiReplace(strText, _
            " ", "&nbsp;", _
            vbCrLf, "<br>")

        ' Closing tags
        If strColor <> vbNullString Then .Add "</font>"
        If blnItalic Then .Add "</i>"
        If blnBold Then .Add "</strong>"

        ' Add line break for HTML
        If blnNextOutputOnNewLine Then m_Console.Add "<br>"
    End With

    ' Allow an update to the screen every second.
    ' (This keeps the aplication from an apparent hang while
    '  running intensive export processes.)
    If m_sngLastUpdate + 1 < Timer Then
        Me.Flush
        m_sngLastUpdate = Timer
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Flush
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Flushes the buffer to the console
'---------------------------------------------------------------------------------------
'
Public Sub Flush()

    ' See if the GUI form is loaded.
    Perf.OperationStart "Console Updates"
    If Not m_RichText Is Nothing Then
        With m_RichText
            m_blnProgressActive = False
            If Not m_Prog Is Nothing Then m_Prog.Hide
            ' Set value, not text to avoid errors with large text strings.
            Echo False
            '.SelStart = Len(.Text & vbNullString)
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = m_Console.RightStr(20000)
            If VerifyFocus(m_RichText) Then .SelStart = 20000
            Echo True
            'Form_frmVCSMain.Repaint
        End With
    Else
        ' Update display in immediate window if not using console
        DoEvents
    End If
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Error
' Author    : Adam Waller
' Date      : 12/4/2020, 7/3/2025 (Josef Poetzl)
' Purpose   : Log an error, and update error level if needed. Reads Err object values.
'           : A critical error will also present a message box with the details.
'---------------------------------------------------------------------------------------
'
Public Sub Error(eLevel As eErrorLevel, strDescription As String, Optional strSource As String, Optional blnIncludeErrorMessage As Boolean = False)

    ErrorInfo NewErrorInfo(eLevel, strDescription, strSource, blnIncludeErrorMessage)

End Sub

Public Sub ErrorInfo(oErrorInfo As clsErrorInfo)

    If oErrorInfo.ErrorLevel <= eelWarning Then
        Exit Sub
    End If

    ' Log all errors, and display on the output screen anything higher than a warning
    Me.Add vbNullString
    Me.Spacer
    Me.Add oErrorInfo.DisplayText, , "red"
    Me.Spacer
    Me.Flush ' Force a screen update.

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetConsole
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set a reference to the rich text box for output display, and the
'           : associated progress bar for logging progress.
'---------------------------------------------------------------------------------------
'
Public Sub SetConsole(txtRichText As TextBox, cProg As clsLblProg)
    Set m_Prog = cProg
    Set m_RichText = txtRichText
    If Not m_RichText Is Nothing Then
        m_RichText.AllowAutoCorrect = False
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReleaseConsole
' Author    : Adam Waller
' Date      : 5/5/2022
' Purpose   : Releases object references to the console form that were created with
'           : the SetConsole sub.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseConsole()
    Set m_RichText = Nothing
    Set m_Prog = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProgressBar
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Pass the Progress Bar reference to this class.
'---------------------------------------------------------------------------------------
'
Public Property Set ProgressBar(cProg As clsLblProg)
    Set m_Prog = cProg
End Property
Public Property Get ProgressBar() As clsLblProg
    Set ProgressBar = m_Prog
End Property


'---------------------------------------------------------------------------------------
' Procedure : ProgMax
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Wrapper to set max value for progress bar.
'---------------------------------------------------------------------------------------
'
Public Property Let ProgMax(lngMaxValue As Long)
    If Not m_Prog Is Nothing Then m_Prog.Max = lngMaxValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Console = New clsConcat
    m_blnProgressActive = False
    m_sngLastUpdate = 0
    m_PadLength = 30
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Pad a string on the right to make it `count` characters long.
'---------------------------------------------------------------------------------------
'
Public Sub PadRight(strText As String, Optional blnNextOutputOnNewLine As Boolean = False, Optional ByVal intCharacters As Integer)
    If intCharacters = 0 Then intCharacters = m_PadLength
    If Len(strText) < intCharacters Then
        Me.Add strText & Space$(intCharacters - Len(strText)), blnNextOutputOnNewLine
    Else
        Me.Add strText, blnNextOutputOnNewLine
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Increment the progress bar/output display
'---------------------------------------------------------------------------------------
'
Public Sub Increment()

    ' Ignore if we are not using the form
    If m_Prog Is Nothing Then Exit Sub
    ' Check the current status.
    Perf.OperationStart "Increment Progress"
    If Not m_blnProgressActive Then
        ' Flush any pending output
        With m_RichText
            Echo False
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = m_Console.RightStr(20000)
            If VerifyFocus(m_RichText) Then .SelStart = 20000
            Echo True
        End With
    End If
    Me.ProgressBar.Increment

    ' Status is now active
    m_blnProgressActive = True
    Perf.OperationEnd

End Sub


'#####################################################
'
' ILogger
'
Private Function ILogger_Init(LogEvents As clsLog) As ILogger
    Set ILogger_Init = Init(LogEvents)
End Function

Public Function Init(LogEvents As clsLog) As clsFormLogger
    Set m_LogEvents = LogEvents
    Set Init = Me
End Function

Private Property Get ILogger_LoggerName() As String
    ILogger_LoggerName = "FormLogger"
End Property


'#####################################################
'
' Log Events
'
Private Sub m_LogEvents_Add(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal strColor As String, ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
    If Not blnPrint Then Exit Sub
    Add strText, blnNextOutputOnNewLine, strColor, blnBold, blnItalic
End Sub

Private Sub m_LogEvents_Clear()
    Clear
End Sub

Private Sub m_LogEvents_ErrorInfo(ByVal oErrorInfo As clsErrorInfo)
    ErrorInfo oErrorInfo
End Sub

Private Sub m_LogEvents_Flush()
    Flush
End Sub

Private Sub m_LogEvents_GetProgressBar(ProgressBarRef As clsLblProg)
    Set ProgressBarRef = ProgressBar
End Sub

Private Sub m_LogEvents_IncrementProgressbar()
    Increment
End Sub

Private Sub m_LogEvents_PadRight(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal intCharacters As Integer)
    If Not blnPrint Then Exit Sub
    PadRight strText, blnNextOutputOnNewLine, intCharacters
End Sub

Private Sub m_LogEvents_ProgMax(ByVal lngMaxValue As Long)
    ProgMax = lngMaxValue
End Sub

Private Sub m_LogEvents_ReleaseConsole()
    ReleaseConsole
End Sub

Private Sub m_LogEvents_SetConsole(ByVal txtRichText As TextBox, ByVal cProg As clsLblProg)
    SetConsole txtRichText, cProg
End Sub

Private Sub m_LogEvents_Spacer(ByVal blnPrint As Boolean)
    If Not blnPrint Then Exit Sub
    Spacer
End Sub

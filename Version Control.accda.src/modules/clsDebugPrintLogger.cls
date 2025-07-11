VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDebugPrintLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Set this to true when logging an operation such as an export or build
' then set back to false after writing the log file. This affects
' how error messages are reported to the user outside of operations.
Public Active As Boolean

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Log As clsConcat      ' Log file output
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
Public Sub Spacer(Optional blnPrint As Boolean = True)
    Me.Add cstrSpacer, blnPrint
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

    If blnNextOutputOnNewLine Then
        ' Create new line
        Debug.Print strText
    Else
        ' Continue next printout on this line.
        Debug.Print strText;
    End If

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
    DoEvents
End Sub


Public Sub ErrorInfo(oErrorInfo As clsErrorInfo)

    If oErrorInfo.ErrorLevel <= eelWarning Then
        Exit Sub
    End If

    ' Build the error message string.
    With New clsConcat

        ' Log all errors, and display on the output screen anything higher than a warning
        Me.Add vbNullString
        Me.Spacer
        Me.Add oErrorInfo.DisplayText

        ' Log the additional error and source details to the log file
        If Err Then .Add "Error ", Err.Number, ": ", Err.Description, " "
        If oErrorInfo.Source <> vbNullString Then .Add "Source: ", oErrorInfo.Source

        Me.Spacer
        Me.Flush ' Force a screen update.
        ' See if we are actively logging an operation
        If Log.Active Then
            ' Show message box for fatal error.
            If oErrorInfo.ErrorLevel = eelCritical Then
                MsgBox2 "Unable to Continue", .GetStr, _
                     "Please review the log file for additional details.", vbCritical
            End If
        Else
            ' Show message on any error level when we are not logging to a file.
            MsgBox2 oErrorInfo.MsgBoxBoldText, oErrorInfo.DisplayText, .GetStr, oErrorInfo.MsgBoxStyle
        End If
    End With

    ' Update error level if higher.
    If Operation.ErrorLevel < oErrorInfo.ErrorLevel Then Operation.ErrorLevel = oErrorInfo.ErrorLevel

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Console = New clsConcat
    Set m_Log = New clsConcat
    m_blnProgressActive = False
    m_sngLastUpdate = 0
    m_PadLength = 30
    Me.Active = False
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

'#####################################################
'
' ILogger
'
Private Function ILogger_Init(LogEvents As clsLog) As ILogger
    Set ILogger_Init = Init(LogEvents)
End Function

Public Function Init(LogEvents As clsLog) As clsDebugPrintLogger
    Set m_LogEvents = LogEvents
    Set Init = Me
End Function

Private Property Get ILogger_LoggerName() As String
    ILogger_LoggerName = "DebugPrintLogger"
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

Private Sub m_LogEvents_PadRight(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal intCharacters As Integer)
    If Not blnPrint Then Exit Sub
    PadRight strText, blnNextOutputOnNewLine, intCharacters
End Sub

Private Sub m_LogEvents_Spacer(ByVal blnPrint As Boolean)
    If Not blnPrint Then Exit Sub
    Spacer
End Sub

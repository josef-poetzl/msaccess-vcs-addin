VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Log As clsConcat      ' Log file output
Private m_PadLength As Integer
Private m_SourcePath As String

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
Public Sub Add(strText As String)

    m_Log.Add strText, vbCrLf

End Sub

Public Sub ErrorInfo(oErrorInfo As clsErrorInfo)

    ' Build the error message string.
    With New clsConcat

        ' Log all errors, and display on the output screen anything higher than a warning
        Me.Add vbNullString
        Me.Spacer
        Me.Add oErrorInfo.DisplayText

        ' Log the additional error and source details to the log file
        If Err Then .Add "Error ", Err.Number, ": ", Err.Description, " "
        If oErrorInfo.Source <> vbNullString Then .Add "Source: ", oErrorInfo.Source
        Me.Add .GetStr
        Me.Spacer

    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveFile
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Saves the current log data to a file.
'---------------------------------------------------------------------------------------
'
Public Sub SaveFile(Optional strAlternatePath As String)
    If Len(strAlternatePath) Then VerifyPath strAlternatePath
    WriteFile m_Log.GetStr, Nz2(strAlternatePath, Me.LogFilePath)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFilePath
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return the (default) full path to the current log file.
'---------------------------------------------------------------------------------------
'
Public Property Get LogFilePath() As String
    Dim strFile As String
    Select Case Operation.OperationType
        Case eotExport: strFile = "Export.log"
        Case eotBuild:  strFile = "Build.log"
        Case eotMerge:  strFile = "Merge.log"
        Case Else:      strFile = "Other.log"
    End Select
    LogFilePath = FSO.BuildPath(Nz2(m_SourcePath, Options.GetExportFolder), strFile)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Log = New clsConcat
    m_PadLength = 30
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Pad a string on the right to make it `count` characters long.
'---------------------------------------------------------------------------------------
'
Public Sub PadRight(strText As String, Optional ByVal intCharacters As Integer)
    If intCharacters = 0 Then intCharacters = m_PadLength
    If Len(strText) < intCharacters Then
        Me.Add strText & Space$(intCharacters - Len(strText))
    Else
        Me.Add strText
    End If
End Sub


'#####################################################
'
' ILogger
'
Private Function ILogger_Init(LogEvents As clsLog) As ILogger
    Set ILogger_Init = Init(LogEvents)
End Function

Public Function Init(LogEvents As clsLog) As clsFileLogger
    Set m_LogEvents = LogEvents
    Set Init = Me
End Function

Private Property Get ILogger_LoggerName() As String
    ILogger_LoggerName = "FileLogger"
End Property


'#####################################################
'
' Log Events
'
Private Sub m_LogEvents_Add(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal strColor As String, ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
    Add strText
End Sub

Private Sub m_LogEvents_Clear()
    Clear
End Sub

Private Sub m_LogEvents_ErrorInfo(ByVal oErrorInfo As clsErrorInfo)
    ErrorInfo oErrorInfo
End Sub

Private Sub m_LogEvents_GetLogFilePath(strFilePath As String)
    strFilePath = LogFilePath
End Sub

Private Sub m_LogEvents_PadRight(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal intCharacters As Integer)
    PadRight strText, intCharacters
End Sub

Private Sub m_LogEvents_SaveFile(ByVal strAlternatePath As String)
    SaveFile strAlternatePath
End Sub

Private Sub m_LogEvents_SourcePath(ByVal strPath As String)
    m_SourcePath = strPath
End Sub

Private Sub m_LogEvents_Spacer(ByVal blnPrint As Boolean)
    Spacer
End Sub

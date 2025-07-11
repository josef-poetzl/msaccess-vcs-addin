VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
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

Public Event Clear()
Public Event Spacer(ByVal blnPrint As Boolean)
Public Event Add(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, _
                 ByVal strColor As String, ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
Public Event PadRight(ByVal strText As String, ByVal blnPrint As Boolean, ByVal blnNextOutputOnNewLine As Boolean, ByVal intCharacters As Integer)
Public Event Flush()
Public Event ErrorInfo(ByVal oErrorInfo As clsErrorInfo)
Public Event IncrementProgressbar()
Public Event GetProgressBar(ByRef ProgressBarRef As clsLblProg)
Public Event ProgMax(ByVal lngMaxValue As Long)
Public Event SetConsole(ByVal txtRichText As TextBox, ByVal cProg As clsLblProg)
Public Event ReleaseConsole()
Public Event SourcePath(ByVal strPath As String)
Public Event SaveFile(ByVal strAlternatePath As String)
Public Event GetLogFilePath(ByRef strFilePath As String)

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the log buffers and release object references
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    RaiseEvent Clear
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
    RaiseEvent Spacer(blnPrint)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add a log file entry.
'---------------------------------------------------------------------------------------
'
Public Sub Add(strText As String, Optional blnPrint As Boolean = True, _
      Optional blnNextOutputOnNewLine As Boolean = True, _
      Optional strColor As String = vbNullString, _
      Optional blnBold As Boolean = False, _
      Optional blnItalic As Boolean = False)

    RaiseEvent Add(strText, blnPrint, blnNextOutputOnNewLine, strColor, blnBold, blnItalic)

End Sub

Public Sub PadRight(strText As String, Optional blnPrint As Boolean = True, Optional blnNextOutputOnNewLine As Boolean = False, Optional ByVal intCharacters As Integer)
    RaiseEvent PadRight(strText, blnPrint, blnNextOutputOnNewLine, intCharacters)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Flush
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Flushes the buffer to the console
'---------------------------------------------------------------------------------------
'
Public Sub Flush()
    RaiseEvent Flush
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

    RaiseEvent ErrorInfo(oErrorInfo)

    ' Update error level if higher.
    If Operation.ErrorLevel < oErrorInfo.ErrorLevel Then Operation.ErrorLevel = oErrorInfo.ErrorLevel

End Sub

Public Property Get ProgressBar() As clsLblProg

    Dim ProgressBarRef As clsLblProg
    RaiseEvent GetProgressBar(ProgressBarRef)
    Set ProgressBar = ProgressBarRef

End Property

Public Property Let ProgMax(lngMaxValue As Long)
    RaiseEvent ProgMax(lngMaxValue)
End Property

Public Property Let SourcePath(strPath As String)
    RaiseEvent SourcePath(strPath)
End Property

'---------------------------------------------------------------------------------------
' Procedure : SaveFile
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Saves the current log data to a file.
'---------------------------------------------------------------------------------------
'
Public Sub SaveFile(Optional strAlternatePath As String)
    RaiseEvent SaveFile(strAlternatePath)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFilePath
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return the (default) full path to the current log file.
'---------------------------------------------------------------------------------------
'
Public Property Get LogFilePath() As String

    Dim strFilePath As String

    RaiseEvent GetLogFilePath(strFilePath)
    LogFilePath = strFilePath

End Property

Public Sub SetConsole(txtRichText As TextBox, cProg As clsLblProg)
    RaiseEvent SetConsole(txtRichText, cProg)
End Sub

Public Sub ReleaseConsole()
    RaiseEvent ReleaseConsole
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Me.Active = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Increment the progress bar/output display
'---------------------------------------------------------------------------------------
'
Public Sub Increment()
    RaiseEvent IncrementProgressbar
End Sub

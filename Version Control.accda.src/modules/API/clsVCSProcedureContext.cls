VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSProcedureContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsVCSProcedureContext
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Restricted view on the running VCS operation for procedures called from
'           : the Run* options (passed with the parameter keyword VCS).
'           : The class holds no reference to clsVersionControl or any other VCS
'           : object. It only knows a snapshot of static values and raises events
'           : for everything else; clsVCSProcedureContextListener handles them on the
'           : VCS side. Without a connected listener (e.g. a reference kept after the
'           : call) writes do nothing and live values return their defaults.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("API")

' Requests to the VCS side (handled by clsVCSProcedureContextListener)
Public Event LogAddRequested(ByVal strText As String, ByVal blnPrint As Boolean, ByVal strColor As String, _
    ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
Public Event LogErrorRequested(ByVal lngLevel As Long, ByVal strDescription As String, ByVal strSource As String)
Public Event ErrorLevelRequested(ByRef lngLevel As Long)
Public Event InteractionModeRequested(ByRef lngMode As Long)

' Snapshot taken when the context is created
Private Type udtThis
    HookName As String
    OperationType As Long
    ExportFolder As String
    DatabasePath As String
    AddInVersion As String
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Set the static values (called by clsProcedureRunner only).
'---------------------------------------------------------------------------------------
'
Friend Sub Init(ByVal strHookName As String, ByVal lngOperationType As Long, _
    ByVal strExportFolder As String, ByVal strDatabasePath As String, ByVal strAddInVersion As String)
    With this
        .HookName = strHookName
        .OperationType = lngOperationType
        .ExportFolder = strExportFolder
        .DatabasePath = strDatabasePath
        .AddInVersion = strAddInVersion
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HookName
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Name of the option that called the procedure (e.g. "RunBeforeExport").
'---------------------------------------------------------------------------------------
'
Public Property Get HookName() As String
    HookName = this.HookName
End Property


'---------------------------------------------------------------------------------------
' Procedure : OperationType
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Type of the running operation (1 = export, 2 = build, 3 = merge, ...).
'---------------------------------------------------------------------------------------
'
Public Property Get OperationType() As Long
    OperationType = this.OperationType
End Property


'---------------------------------------------------------------------------------------
' Procedure : ExportFolder
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Source folder of the current database.
'---------------------------------------------------------------------------------------
'
Public Property Get ExportFolder() As String
    ExportFolder = this.ExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : DatabasePath
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Full path of the current database.
'---------------------------------------------------------------------------------------
'
Public Property Get DatabasePath() As String
    DatabasePath = this.DatabasePath
End Property


'---------------------------------------------------------------------------------------
' Procedure : AddInVersion
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Version of the VCS add-in.
'---------------------------------------------------------------------------------------
'
Public Property Get AddInVersion() As String
    AddInVersion = this.AddInVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : ErrorLevel
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Current error level of the operation (0 = none, 1 = warning, 2 = info,
'           : 3 = alert, 4 = error, 5 = critical). 0 without a connected listener.
'---------------------------------------------------------------------------------------
'
Public Property Get ErrorLevel() As Long
    Dim lngLevel As Long
    RaiseEvent ErrorLevelRequested(lngLevel)
    ErrorLevel = lngLevel
End Property


'---------------------------------------------------------------------------------------
' Procedure : InteractionMode
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : 0 = normal, 1 = silent (no dialogs). 0 without a connected listener.
'---------------------------------------------------------------------------------------
'
Public Property Get InteractionMode() As Long
    Dim lngMode As Long
    RaiseEvent InteractionModeRequested(lngMode)
    InteractionMode = lngMode
End Property


'---------------------------------------------------------------------------------------
' Procedure : AddLog
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Add a line to the VCS log (and to the console if blnPrint).
'---------------------------------------------------------------------------------------
'
Public Sub AddLog(ByVal strText As String, Optional ByVal blnPrint As Boolean = True, _
    Optional ByVal strColor As String = vbNullString, Optional ByVal blnBold As Boolean = False, _
    Optional ByVal blnItalic As Boolean = False)
    RaiseEvent LogAddRequested(strText, blnPrint, strColor, blnBold, blnItalic)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogError
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Log a message with an error level (1 = warning, 2 = info, 3 = alert,
'           : 4 = error, 5 = critical, which cancels the operation).
'---------------------------------------------------------------------------------------
'
Public Sub LogError(ByVal lngLevel As Long, ByVal strDescription As String, _
    Optional ByVal strSource As String = vbNullString)
    RaiseEvent LogErrorRequested(lngLevel, strDescription, strSource)
End Sub

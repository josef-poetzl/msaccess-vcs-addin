VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSProcedureContextListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVCSProcedureContextListener
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Handles the events of a clsVCSProcedureContext on the VCS side and
'           : forwards them to Log and Operation. Holds no clsVersionControl instance.
'           : Detach disconnects the context, so a reference that an external
'           : procedure keeps afterwards reaches nothing.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("API")

Private Const ModuleName As String = "clsVCSProcedureContextListener"

Private WithEvents m_Context As clsVCSProcedureContext
Attribute m_Context.VB_VarHelpID = -1


'---------------------------------------------------------------------------------------
' Procedure : Attach
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Start handling the events of oContext.
'---------------------------------------------------------------------------------------
'
Friend Sub Attach(ByVal oContext As clsVCSProcedureContext)
    Set m_Context = oContext
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Detach
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Stop handling events and release the context.
'---------------------------------------------------------------------------------------
'
Friend Sub Detach()
    Set m_Context = Nothing
End Sub


Private Sub Class_Terminate()
    Detach
End Sub


Private Sub m_Context_LogAddRequested(ByVal strText As String, ByVal blnPrint As Boolean, _
    ByVal strColor As String, ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
    ForwardLogAdd strText, blnPrint, strColor, blnBold, blnItalic
End Sub


Private Sub m_Context_LogErrorRequested(ByVal lngLevel As Long, ByVal strDescription As String, _
    ByVal strSource As String)
    ForwardLogError lngLevel, strDescription, strSource
End Sub


Private Sub m_Context_ErrorLevelRequested(ByRef lngLevel As Long)
    lngLevel = Operation.ErrorLevel
End Sub


Private Sub m_Context_InteractionModeRequested(ByRef lngMode As Long)
    lngMode = Operation.InteractionMode
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ForwardLogAdd
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Add the external text to the log.
'---------------------------------------------------------------------------------------
'
Private Sub ForwardLogAdd(ByVal strText As String, ByVal blnPrint As Boolean, _
    ByVal strColor As String, ByVal blnBold As Boolean, ByVal blnItalic As Boolean)
    Log.Add strText, blnPrint, , strColor, blnBold, blnItalic
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ForwardLogError
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Log the external message with a valid error level. An unknown level is
'           : logged as eelAlert; a missing source names the calling option.
'---------------------------------------------------------------------------------------
'
Private Sub ForwardLogError(ByVal lngLevel As Long, ByVal strDescription As String, _
    ByVal strSource As String)

    Dim eLevel As eErrorLevel

    If lngLevel < eelWarning Or lngLevel > eelCritical Then
        eLevel = eelAlert
    Else
        eLevel = lngLevel
    End If
    If Len(strSource) = 0 Then strSource = m_Context.HookName

    Log.Error eLevel, strDescription, strSource

End Sub

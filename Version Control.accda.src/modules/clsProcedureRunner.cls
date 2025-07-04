VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcedureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Const ModuleName As String = "clsProcedureRunner"

'---------------------------------------------------------------------------------------
' Procedure : RunInCurrentProject
' Author    : Adam Waller
' Date      : 4/22/2020
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunProcedureInCurrentProject(strSubName As String)

    Dim strSub As String
    Dim strCmd As String

    ' Don't need the parentheses after the sub name
    strSub = Replace(strSubName, "()", vbNullString)

    ' Make sure we are not trying to run a function with arguments
    If InStr(strSub, "(") > 0 Then
        MsgBox2 T("Unable to Run Command"), _
            T("Parameters are not supported for this command."), _
            T("If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of {0}.", var0:=strSubName), vbExclamation
        Exit Sub
    End If

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSub) Then
        Log.Error eelError, T("The procedure ""{0}"" not found.", var0:=strSub), ModuleName & ".RunSubInCurrentProject"
        Log.Add T("The procedure must be declared as public in a standard module."), False
        Exit Sub
    End If

    ' Build call syntax
    If CurrentVBProject.Name = PROJECT_NAME Then
        ' use full path
        ' Example: Run "c:\full\path\Version Control.SubName"
        With CurrentProject
            strCmd = .Path & PathSep & FSO.GetBaseName(.Name) & "." & strSub
        End With
    Else
        ' use library name
        ' Example: Run "[VBProject].SubName"
        strCmd = "[" & CurrentVBProject.Name & "]." & strSub
    End If

    ' Log any outstanding errors
    LogUnhandledErrors

    ' Stage the current operation, and run the sub
    Operation.Stage
    Perf.OperationStart T("Run {0}", , , , strSub)

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    Application.Run strCmd
    Perf.OperationEnd
    Operation.Restore

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strSub), ModuleName & ".RunSubInCurrentProject"

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcedureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const ModuleName As String = "clsProcedureRunner"

'---------------------------------------------------------------------------------------
' Procedure : RunExternalProcedure
' Author    : Josef Poetzl
' Date      : 7/6/2025
' Purpose   : Runs one or more external procedures by name, which is specified in the
'           : RunAfterBuild, RunAfterMerge etc. options.
'           : The procedure name can be a single procedure or multiple procedures
'           : separated by a colon.
'---------------------------------------------------------------------------------------
'
Public Sub RunExternalProcedure(ByVal strRunProcedureOptionValue As String, Optional ByVal VcsRef As clsVersionControl = Nothing)

    Dim arrProcedures() As String
    Dim i As Long
    Dim blnRan As Boolean
    Dim strProcedureCallExpression As String

    arrProcedures = Split(strRunProcedureOptionValue, ":")
    ' TODO: Split with regexp so that string parameters can contain ":"

    For i = LBound(arrProcedures) To UBound(arrProcedures)
        blnRan = False
        strProcedureCallExpression = Trim(arrProcedures(i))
        Log.Add T("Running {0}...", var0:=strProcedureCallExpression)
        If Not TryRunAddInProcedure(strProcedureCallExpression, VcsRef, blnRan) Then
            RunProcInCurrentProject strProcedureCallExpression, blnRan, True, VcsRef, True
        End If
        If Not blnRan Then
            Log.Error eelWarning, T("{0} not finished", var0:=strProcedureCallExpression), ModuleName & ".RunExternalProcedure"
        End If
    Next

End Sub

Private Function TryRunAddInProcedure(ByVal strProcedureCallExpression As String, ByVal VcsRef As clsVersionControl, ByRef blnRan As Boolean) As Boolean

    Dim strAddInFile As String
    Dim arrParams() As Variant

    strProcedureCallExpression = Replace(strProcedureCallExpression, "%addins%", Environ$("appdata") & "\Microsoft\AddIns", , , vbTextCompare)
    strProcedureCallExpression = Replace(strProcedureCallExpression, "%appdata%", Environ("appdata"), , , vbTextCompare)

    strAddInFile = Left(strProcedureCallExpression, InStrRev(strProcedureCallExpression, ".")) & "accda"
    If Len(VBA.Dir(strAddInFile)) = 0 Then
        TryRunAddInProcedure = False
        Exit Function
    End If

    TryRunAddInProcedure = True

    If Not PrepareProcedureCallExpressionAndParams(strProcedureCallExpression, arrParams, VcsRef) Then
        Exit Function
    End If

    CallApplicationRun strProcedureCallExpression, arrParams, blnRan, True, True


End Function


'---------------------------------------------------------------------------------------
' Procedure : RunProcInCurrentProject
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Run a procedure in the current project and return whatever it returns.
'           : Identical to RunSubInCurrentProject except that the return value is
'           : surfaced, which is what a validation hook needs. blnRan distinguishes
'           : "ran and returned Empty" from "never ran" (missing, or rejected for
'           : taking parameters) -- a caller that treats a hook failure as a build
'           : failure must not read a refusal as a False return.
'           :
'           : blnWantResult exists so RunSubInCurrentProject keeps calling
'           : Application.Run as a statement. Assigning its result to a Variant
'           : raises "object variable not set" when the procedure returns an object,
'           : which would newly break existing RunAfter* hooks that happen to do so.
'---------------------------------------------------------------------------------------
'
Public Function RunProcInCurrentProject(strSubName As String, _
    Optional ByRef blnRan As Boolean, _
    Optional ByVal blnWantResult As Boolean = True, _
    Optional ByVal VcsRef As clsVersionControl = Nothing, _
    Optional ByVal UseExternalLogReturn As Boolean = False) As Variant

    Dim strSub As String
    Dim strCmd As String
    Dim arrParams() As Variant

    blnRan = False

    strSub = strSubName

    If Not PrepareProcedureCallExpressionAndParams(strSub, arrParams, VcsRef) Then
        Exit Function
    End If

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSub) Then
        Log.Error eelError, T("The procedure ""{0}"" not found.", var0:=strSub), ModuleName & ".RunSubInCurrentProject"
        Log.Add T("The procedure must be declared as public in a standard module."), False
        Exit Function
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

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    If blnWantResult Then
        RunProcInCurrentProject = CallApplicationRun(strCmd, arrParams, blnRan, blnWantResult, UseExternalLogReturn)
    Else
        CallApplicationRun strCmd, arrParams, blnRan, blnWantResult
    End If

End Function

Friend Function CallApplicationRun(ByVal strProcedure As String, ByRef arrParams() As Variant, ByRef blnRan As Boolean, ByVal blnWantResult As Boolean, _
                          Optional ByVal UseExternalLogReturn As Boolean = False) As Variant

' if UseExternalLogReturn = True:
' What could a generally usable interface look like?
'
' * Public Function ProcedureName(ByRef ReturnMessage As String) as Boolean
' * Public Function ProcedureName(ByRef ReturnMessage As String) as Long ' ... = eErrorLevel .. -1 for all ok?
' * Public Function ProcedureName() as String ... Returns:
'                                      "Error: ErrorMessage"   => Error log
'                                   or "Warning: Warning Message" => displayed Warning log
'                                   or vbNullstring ... show nothing, all success
'
' used interface:
' * Public Function ProcedureName() as Variant ... Returns:
'                                      "Error: ErrorMessage"   => Error log
'                                   or "Warning: Warning Message" => displayed Warning log
'                                   or ... (see LogErrorMessage)
'                                   or vbNullstring ... show nothing, all success
'                                   or True/False .. True = success, false = failed


    Dim lngParamCnt As Long
    Dim MinIndex As Long
    Dim ProcReturnValue As Variant
    Dim cPause As clsOperationPause

    If Not CBool(Not Not arrParams) Then
        lngParamCnt = 0
    Else
        MinIndex = LBound(arrParams)
        lngParamCnt = UBound(arrParams) - MinIndex + 1
    End If

    If lngParamCnt > 5 Then
        Err.Raise vbObjectError, ModuleName & ".CallApplicationRun", "Only up to 5 parameters are supported"
    End If

    ' Log any outstanding errors
    LogUnhandledErrors

    ' Pause the current operation while user code runs. The user's procedure is foreign
    ' code: it may raise, open forms, or take its own time, none of which should look
    ' like our operation still running. ResumePause ends the pause on the success path;
    ' Class_Terminate on cPause is the fallback if Application.Run raises.
    Set cPause = Operation.TryPause()
    Perf.OperationStart T("Run {0}", , , , strProcedure)

    If blnWantResult Or UseExternalLogReturn Then
        Select Case lngParamCnt
            Case 0
                ProcReturnValue = Application.Run(strProcedure)
            Case 1
                ProcReturnValue = Application.Run(strProcedure, arrParams(MinIndex))
            Case 2
                ProcReturnValue = Application.Run(strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1))
            Case 3
                ProcReturnValue = Application.Run(strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2))
            Case 4
                ProcReturnValue = Application.Run(strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2), arrParams(MinIndex + 3))
            Case 5
                ProcReturnValue = Application.Run(strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2), arrParams(MinIndex + 3), arrParams(MinIndex + 4))
        End Select
    Else
        Select Case lngParamCnt
            Case 0
                Application.Run (strProcedure)
            Case 1
                Application.Run strProcedure, arrParams(MinIndex)
            Case 2
                Application.Run strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1)
            Case 3
                Application.Run strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2)
            Case 4
                Application.Run strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2), arrParams(MinIndex + 3)
            Case 5
                Application.Run strProcedure, arrParams(MinIndex), arrParams(MinIndex + 1), arrParams(MinIndex + 2), arrParams(MinIndex + 3), arrParams(MinIndex + 4)
        End Select
    End If

    blnRan = True
    Perf.OperationEnd
    If Not cPause Is Nothing Then cPause.ResumePause

    If UseExternalLogReturn Then
        If VarType(ProcReturnValue) = vbString Then
            LogErrorMessage ProcReturnValue, GetProcedureNameFromPath(strProcedure)
        ElseIf VarType(ProcReturnValue) = vbBoolean Then
            If Not ProcReturnValue Then ' eelCritical => Cancel
                Log.Error eelCritical, GetProcedureNameFromPath(strProcedure) & " failed (return False)", ModuleName & ".CallApplicationRun"
            End If
        End If
    End If

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strProcedure), ModuleName & ".CallApplicationRun"

    CallApplicationRun = ProcReturnValue

End Function

Friend Function PrepareProcedureCallExpressionAndParams(ByRef strProcedureCallExpression As String, ByRef arrParams() As Variant, ByVal VcsRef As clsVersionControl) As Boolean

    ' Don't need the parentheses after the sub name
    strProcedureCallExpression = Replace(strProcedureCallExpression, "()", vbNullString)

    If InStr(strProcedureCallExpression, "(") > 0 Then
        If Not TryGetProcedureNamenAndParams(strProcedureCallExpression, arrParams, VcsRef) Then
            MsgBox2 T("Unable to Run Command"), _
                T("Unable to parse parameters from command."), _
                T("Please check the parameter syntax. Parameters should be separated by commas and" & vbCrLf & _
                "quoted strings should use double quotes." & vbCrLf & _
                "Example: {0}", var0:="TestProc(""param1"", param2, ""param3"")"), vbExclamation
            PrepareProcedureCallExpressionAndParams = False
            Exit Function
        End If
    End If

    PrepareProcedureCallExpressionAndParams = True

End Function


Friend Function TryGetProcedureNamenAndParams(ByRef strProcedure As String, ByRef arrParams() As Variant, ByVal VcsRef As clsVersionControl) As Boolean

    Dim strProcName As String
    Dim strParams As String
    Dim lngParamDeclLeft As Long
    Dim lngParamDeclRight As Long

    Dim arrParamList() As Variant
    Dim i As Long
    Dim varParam As Variant

    lngParamDeclLeft = InStr(1, strProcedure, "(")
    If lngParamDeclLeft = 0 Then
        TryGetProcedureNamenAndParams = False
        Exit Function
    End If

    lngParamDeclRight = InStrRev(strProcedure, ")")
    If lngParamDeclRight = 0 Then
        TryGetProcedureNamenAndParams = False
        Exit Function
    End If

    strProcName = Trim(Left(strProcedure, lngParamDeclLeft - 1))
    strParams = Trim(Mid(strProcedure, lngParamDeclLeft + 1, lngParamDeclRight - lngParamDeclLeft - 1))

    ' If no parameters, return success with empty array
    If Len(strParams) = 0 Then
        TryGetProcedureNamenAndParams = True
        Exit Function
    End If

    arrParamList = ParseParameterString(strParams)

    If Not CBool(Not Not arrParamList) Then
        TryGetProcedureNamenAndParams = False
        Exit Function
    End If

    ReDim arrParams(0 To UBound(arrParamList))

    ' Clean up each parameter (remove quotes and trim whitespace)
    For i = 0 To UBound(arrParamList)
        varParam = Trim(arrParamList(i))
        If varParam = "Nothing" Then
            Set arrParams(i) = Nothing
        ElseIf varParam = "VCS" Then
            Set arrParams(i) = VcsRef
        Else
            If varParam = "Null" Then
                varParam = Null
            ElseIf Len(varParam) >= 2 Then
                If (Left(varParam, 1) = """" And Right(varParam, 1) = """") Then
                    varParam = CStr(Mid(varParam, 2, Len(varParam) - 2))
                End If
            End If
            arrParams(i) = varParam
        End If
    Next i

    strProcedure = strProcName
    TryGetProcedureNamenAndParams = True

End Function

Friend Function ParseParameterString(ByVal strParams As String) As Variant()

    Dim objRegex As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim objMatch As VBScript_RegExp_55.Match

    Dim arrResult() As Variant
    Dim i As Long

    ' Create regex object
    Set objRegex = New VBScript_RegExp_55.RegExp

    ' Pattern explanation (GitHub CoPilot):
    ' \s* - optional leading whitespace
    ' ( - start capture group
    '   "[^"]*" - double-quoted strings (including commas inside)
    '   | - or
    '   '[^']*' - single-quoted strings (including commas inside)
    '   | - or
    '   [^,]+ - unquoted parameters (non-comma characters)
    ' ) - end capture group
    ' \s* - optional trailing whitespace
    ' (?:,|$) - followed by comma or end of string (non-capturing)
    objRegex.Pattern = "\s*(""[^""]*""|'[^']*'|[^,]+)\s*(?:,|$)"
    objRegex.Global = True
    objRegex.IgnoreCase = True

    ' Find all matches
    Set objMatches = objRegex.Execute(strParams)

    ' Convert matches to array
    If objMatches.Count > 0 Then
        ReDim arrResult(0 To objMatches.Count - 1)
        For i = 0 To objMatches.Count - 1
            ' Use SubMatches(0) to get the captured group content (parameter without commas)
            arrResult(i) = objMatches(i).SubMatches(0)
        Next i
    Else
        ' No matches found - return unallocated array
    End If

    ParseParameterString = arrResult

End Function

Private Sub LogErrorMessage(ByVal strErrorMessage As String, ByVal strErrorMessageSource As String)

    Dim lngErrorLevel As eErrorLevel
    Dim lngErrorLevelEndPos As Long

    lngErrorLevelEndPos = InStr(1, strErrorMessage, ":")
    If lngErrorLevelEndPos > 1 Then
        Select Case Trim(Left(strErrorMessage, lngErrorLevelEndPos - 1))
            Case "Error"
                lngErrorLevel = eelError
            Case "Warning", "Alert", "Failed"
                lngErrorLevel = eelAlert
            Case "Critical", "FATAL"
                lngErrorLevel = eelCritical
            Case "Note", "Success", "Info"
                lngErrorLevel = eelInfo
            Case "Log"
                lngErrorLevel = eelWarning
            Case Else
                lngErrorLevel = eelAlert
                lngErrorLevelEndPos = 0 ' don't remove String before ":"
        End Select
        If lngErrorLevelEndPos > 0 Then
            strErrorMessage = Trim(Mid(strErrorMessage, lngErrorLevelEndPos + 1))
        End If
    Else
        lngErrorLevel = eelAlert
    End If

    Log.ErrorInfo NewErrorInfo(lngErrorLevel, strErrorMessage, strErrorMessageSource, False)

End Sub

Private Function GetProcedureNameFromPath(ByVal strFullProcedureName As String) As String

    Dim lngCutPos As Long

    lngCutPos = InStrRev(Replace(strFullProcedureName, "!", "\"), "\")
    If lngCutPos > 0 Then
        GetProcedureNameFromPath = Mid(strFullProcedureName, lngCutPos + 1)
    Else
        GetProcedureNameFromPath = strFullProcedureName
    End If

End Function

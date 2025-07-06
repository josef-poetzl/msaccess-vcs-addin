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

    arrProcedures = Split(strRunProcedureOptionValue, ":")
    ' TODO: Split with regexp so that string parameters can contain ":"

    For i = LBound(arrProcedures) To UBound(arrProcedures)
        Log.Add T("Running {0}...", var0:=arrProcedures(i))
        RunProcedureInCurrentProject Trim(arrProcedures(i)), VcsRef
    Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunProcedureInCurrentProject
' Author    : Adam Waller, Josef Poetzl
' Date      : 4/22/2020, 7/4/2025
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunProcedureInCurrentProject(strSubName As String, Optional ByVal VcsRef As clsVersionControl = Nothing)

    Dim strSub As String
    Dim strCmd As String
    Dim arrParams() As Variant

    ' Don't need the parentheses after the sub name
    strSub = Replace(strSubName, "()", vbNullString)

    ' Make sure we are not trying to run a function with arguments
    If InStr(strSub, "(") > 0 Then
        If Not TryGetProcedureNamenAndParams(strSub, arrParams, VcsRef) Then
            MsgBox2 T("Unable to Run Command"), _
                T("Unable to parse parameters from command."), _
                T("Please check the parameter syntax. Parameters should be separated by commas and" & vbCrLf & _
                "quoted strings should use double quotes." & vbCrLf & _
                "Example: {0}", var0:="TestProc(""param1"", param2, ""param3"")"), vbExclamation
            Exit Sub
        Else
            ' Extract the procedure name from the command with parameters
            strSub = Trim(Left(strSub, InStr(strSub, "(") - 1))
        End If
    End If

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSub) Then
        Log.Error eelError, T("The procedure ""{0}"" not found.", var0:=strSub), ModuleName & ".RunProcedureInCurrentProject"
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

    CallApplicationRun strCmd, arrParams
    Perf.OperationEnd
    Operation.Restore

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strSub), ModuleName & ".RunProcedureInCurrentProject"

End Sub

Friend Sub CallApplicationRun(ByVal strProcedure As String, ByRef arrParams() As Variant)

    Dim lngParamCnt As Long
    Dim MinIndex As Long

    If Not CBool(Not Not arrParams) Then
        lngParamCnt = 0
    Else
        MinIndex = LBound(arrParams)
        lngParamCnt = UBound(arrParams) - MinIndex + 1
    End If

    Select Case lngParamCnt
        Case 0
            Application.Run strProcedure
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

        Case Else
            Err.Raise vbObjectError, ModuleName & ".CallApplicationRun", "Only up to 5 parameters are supported"
    End Select

End Sub

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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcedureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller, Josef Pötzl
' Date      : 4/23/2020
' Purpose   : Encapsulates the execution of procedures within the current Microsoft Access project.
'           : The class provides a dedicated and extensible abstraction for procedure execution,
'           : keeping this functionality separate from database-related concerns.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

Private Const ModuleName As String = "clsProcedureRunner"

' Application.Run accepts up to 30 arguments; CallApplicationRun maps this many.
Private Const MAX_PARAM_COUNT As Long = 5

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
    Optional ByVal blnWantResult As Boolean = True) As Variant

    Dim strSub As String
    Dim strCmd As String
    Dim arrParams() As Variant

    blnRan = False

    ' Split the call expression into procedure name and parameter values
    strSub = strSubName
    If Not PrepareProcedureCallExpressionAndParams(strSub, arrParams) Then Exit Function

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
        RunProcInCurrentProject = CallApplicationRun(strCmd, arrParams, blnRan, True)
    Else
        CallApplicationRun strCmd, arrParams, blnRan, False
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CallApplicationRun
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Run the procedure via Application.Run with up to MAX_PARAM_COUNT
'           : parameters. With blnWantResult = False Application.Run is called as a
'           : statement (see RunProcInCurrentProject for the reason).
'---------------------------------------------------------------------------------------
'
Friend Function CallApplicationRun(ByVal strCmd As String, ByRef arrParams() As Variant, _
    ByRef blnRan As Boolean, ByVal blnWantResult As Boolean) As Variant

    Dim lngParamCount As Long
    Dim lngMin As Long
    Dim strProcName As String
    Dim cPause As clsOperationPause

    blnRan = False
    strProcName = GetProcedureNameFromCmd(strCmd)

    If Not IsEmptyArray(arrParams) Then
        lngMin = LBound(arrParams)
        lngParamCount = UBound(arrParams) - lngMin + 1
    End If

    If lngParamCount > MAX_PARAM_COUNT Then
        Log.Error eelError, T("Only up to {0} parameters are supported ({1}).", _
            var0:=MAX_PARAM_COUNT, var1:=strProcName), ModuleName & ".CallApplicationRun"
        Exit Function
    End If

    ' Log any outstanding errors
    LogUnhandledErrors

    ' Pause the current operation while user code runs. The user's procedure is foreign
    ' code: it may raise, open forms, or take its own time, none of which should look
    ' like our operation still running. ResumePause ends the pause on the success path;
    ' Class_Terminate on cPause is the fallback if Application.Run raises.
    Set cPause = Operation.TryPause()
    Perf.OperationStart T("Run {0}", , , , strProcName)

    If blnWantResult Then
        Select Case lngParamCount
            Case 0: CallApplicationRun = Application.Run(strCmd)
            Case 1: CallApplicationRun = Application.Run(strCmd, arrParams(lngMin))
            Case 2: CallApplicationRun = Application.Run(strCmd, arrParams(lngMin), arrParams(lngMin + 1))
            Case 3: CallApplicationRun = Application.Run(strCmd, arrParams(lngMin), arrParams(lngMin + 1), _
                        arrParams(lngMin + 2))
            Case 4: CallApplicationRun = Application.Run(strCmd, arrParams(lngMin), arrParams(lngMin + 1), _
                        arrParams(lngMin + 2), arrParams(lngMin + 3))
            Case 5: CallApplicationRun = Application.Run(strCmd, arrParams(lngMin), arrParams(lngMin + 1), _
                        arrParams(lngMin + 2), arrParams(lngMin + 3), arrParams(lngMin + 4))
        End Select
    Else
        Select Case lngParamCount
            Case 0: Application.Run strCmd
            Case 1: Application.Run strCmd, arrParams(lngMin)
            Case 2: Application.Run strCmd, arrParams(lngMin), arrParams(lngMin + 1)
            Case 3: Application.Run strCmd, arrParams(lngMin), arrParams(lngMin + 1), arrParams(lngMin + 2)
            Case 4: Application.Run strCmd, arrParams(lngMin), arrParams(lngMin + 1), arrParams(lngMin + 2), _
                        arrParams(lngMin + 3)
            Case 5: Application.Run strCmd, arrParams(lngMin), arrParams(lngMin + 1), arrParams(lngMin + 2), _
                        arrParams(lngMin + 3), arrParams(lngMin + 4)
        End Select
    End If
    blnRan = True
    Perf.OperationEnd
    If Not cPause Is Nothing Then cPause.ResumePause

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strProcName), ModuleName & ".CallApplicationRun"

End Function


'---------------------------------------------------------------------------------------
' Procedure : PrepareProcedureCallExpressionAndParams
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Reduce a call expression such as MyProc("a", 1) to the procedure name
'           : and return the parsed parameter values. Reports a syntax error in the
'           : log and, when a dialog is allowed, in a message box (MsgBox2 only logs
'           : in silent mode, so unattended runs are not blocked).
'---------------------------------------------------------------------------------------
'
Friend Function PrepareProcedureCallExpressionAndParams(ByRef strProcedureCallExpression As String, _
    ByRef arrParams() As Variant) As Boolean

    Dim strOriginal As String

    strOriginal = strProcedureCallExpression

    ' Don't need the parentheses after the sub name
    strProcedureCallExpression = Trim$(Replace(strProcedureCallExpression, "()", vbNullString))

    If InStr(strProcedureCallExpression, "(") > 0 Then
        If Not TryGetProcedureNameAndParams(strProcedureCallExpression, arrParams) Then
            Log.Error eelError, T("Unable to parse parameters from command: {0}", var0:=strOriginal), _
                ModuleName & ".PrepareProcedureCallExpressionAndParams"
            MsgBox2 T("Unable to Run Command"), _
                T("Unable to parse parameters from command."), _
                T("Please check the parameter syntax. Parameters should be separated by commas and" & vbCrLf & _
                "strings should be enclosed in quotes." & vbCrLf & _
                "Example: {0}", var0:="TestProc(""param1"", param2, ""param3"")"), vbExclamation
            Exit Function
        End If
    End If

    PrepareProcedureCallExpressionAndParams = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : TryGetProcedureNameAndParams
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Split "ProcName(param1, param2, ...)" into the procedure name (returned
'           : in strProcedure) and the parameter values.
'           : Supported values: "text" or 'text' (passed without quotes, a doubled
'           : quote inside stands for one quote), Null, Nothing, and unquoted values
'           : (passed as string, converted by VBA).
'---------------------------------------------------------------------------------------
'
Friend Function TryGetProcedureNameAndParams(ByRef strProcedure As String, _
    ByRef arrParams() As Variant) As Boolean

    Dim lngLeft As Long
    Dim lngRight As Long
    Dim strProcName As String
    Dim strParams As String
    Dim varTokens As Variant
    Dim lngIndex As Long

    lngLeft = InStr(1, strProcedure, "(")
    lngRight = InStrRev(strProcedure, ")")
    If lngLeft = 0 Or lngRight < lngLeft Then Exit Function

    ' Nothing but whitespace may follow the closing parenthesis
    If Len(Trim$(Mid$(strProcedure, lngRight + 1))) > 0 Then Exit Function

    strProcName = Trim$(Left$(strProcedure, lngLeft - 1))
    If Len(strProcName) = 0 Then Exit Function

    strParams = Trim$(Mid$(strProcedure, lngLeft + 1, lngRight - lngLeft - 1))
    If Len(strParams) > 0 Then
        If Not TryParseParameterString(strParams, varTokens) Then Exit Function
        ReDim arrParams(0 To UBound(varTokens))
        For lngIndex = 0 To UBound(varTokens)
            SetParameterValue arrParams, lngIndex, CStr(varTokens(lngIndex))
        Next lngIndex
    End If

    strProcedure = strProcName
    TryGetProcedureNameAndParams = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : TryParseParameterString
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Split a comma separated parameter list into trimmed tokens. Commas
'           : inside quoted strings do not separate. Returns False for empty or
'           : unbalanced tokens.
'---------------------------------------------------------------------------------------
'
Friend Function TryParseParameterString(ByVal strParams As String, ByRef varTokens As Variant) As Boolean

    Dim oRegEx As VBScript_RegExp_55.RegExp
    Dim oMatches As VBScript_RegExp_55.MatchCollection
    Dim strTokens() As String
    Dim lngIndex As Long
    Dim lngMatchedLength As Long

    Set oRegEx = New VBScript_RegExp_55.RegExp
    With oRegEx
        ' Leading whitespace, then a double-quoted string, a single-quoted string
        ' (a doubled quote inside escapes the quote), or any run of non-comma
        ' characters, then a comma or the end of the string.
        .Pattern = "\s*(""(?:[^""]|"""")*""|'(?:[^']|'')*'|[^,]+)\s*(,|$)"
        .Global = True
        Set oMatches = .Execute(strParams)
    End With

    If oMatches.Count = 0 Then Exit Function

    ReDim strTokens(0 To oMatches.Count - 1)
    For lngIndex = 0 To oMatches.Count - 1
        lngMatchedLength = lngMatchedLength + oMatches(lngIndex).Length
        strTokens(lngIndex) = Trim$(oMatches(lngIndex).SubMatches(0))
        If Not IsValidParameterToken(strTokens(lngIndex)) Then Exit Function
    Next lngIndex

    ' Skipped characters (such as an empty value between two commas) mean a syntax error
    If lngMatchedLength <> Len(strParams) Then Exit Function

    ' A trailing comma leaves out the last value
    If Right$(strParams, 1) = "," Then Exit Function

    varTokens = strTokens
    TryParseParameterString = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsValidParameterToken
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : A token is valid if it is not empty and a quoted string is closed at
'           : its end and contains its quote character only doubled.
'---------------------------------------------------------------------------------------
'
Private Function IsValidParameterToken(ByVal strToken As String) As Boolean

    Dim strQuote As String
    Dim strInner As String

    If Len(strToken) = 0 Then Exit Function

    strQuote = Left$(strToken, 1)
    If strQuote = """" Or strQuote = "'" Then
        If Len(strToken) < 2 Then Exit Function
        If Right$(strToken, 1) <> strQuote Then Exit Function
        strInner = Mid$(strToken, 2, Len(strToken) - 2)
        If InStr(1, Replace(strInner, strQuote & strQuote, vbNullString), strQuote) > 0 Then Exit Function
    End If

    IsValidParameterToken = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetParameterValue
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Convert a parameter token into the value passed to Application.Run.
'---------------------------------------------------------------------------------------
'
Private Sub SetParameterValue(ByRef arrParams() As Variant, ByVal lngIndex As Long, ByVal strToken As String)

    Dim strQuote As String

    strQuote = Left$(strToken, 1)

    If strQuote = """" Or strQuote = "'" Then
        arrParams(lngIndex) = Replace(Mid$(strToken, 2, Len(strToken) - 2), strQuote & strQuote, strQuote)
    ElseIf StrComp(strToken, "Nothing", vbTextCompare) = 0 Then
        Set arrParams(lngIndex) = Nothing
    ElseIf StrComp(strToken, "Null", vbTextCompare) = 0 Then
        arrParams(lngIndex) = Null
    Else
        arrParams(lngIndex) = strToken
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetProcedureNameFromCmd
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Return the bare procedure name of a call command such as
'           : "[Project].MyProc" or "C:\Path\Project.MyProc".
'---------------------------------------------------------------------------------------
'
Private Function GetProcedureNameFromCmd(ByVal strCmd As String) As String
    GetProcedureNameFromCmd = Mid$(strCmd, InStrRev(strCmd, ".") + 1)
End Function

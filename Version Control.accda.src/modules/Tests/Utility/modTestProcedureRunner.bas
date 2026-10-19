Attribute VB_Name = "modTestProcedureRunner"
'---------------------------------------------------------------------------------------
' Module    : modTestProcedureRunner
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Unit tests for parsing procedure call expressions in clsProcedureRunner.
'           : (No procedure is executed.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")


Public Sub TestProcedureCallWithoutParams()

    Dim strExpr As String
    Dim arrParams() As Variant

    strExpr = "MyProc"
    TestAssert ParseCall(strExpr, arrParams), "plain name is accepted"
    TestAssert strExpr = "MyProc", "name unchanged"
    TestAssert IsEmptyArray(arrParams), "no parameters"

    strExpr = "MyProc()"
    TestAssert ParseCall(strExpr, arrParams), "empty parentheses are accepted"
    TestAssert strExpr = "MyProc", "parentheses removed"
    TestAssert IsEmptyArray(arrParams), "no parameters for ()"

    strExpr = "MyProc(  )"
    TestAssert ParseCall(strExpr, arrParams), "whitespace in parentheses is accepted"
    TestAssert strExpr = "MyProc", "name without parentheses"

End Sub


Public Sub TestProcedureCallWithParams()

    Dim strExpr As String
    Dim arrParams() As Variant

    strExpr = "MyProc(""a, b"", 'c', 12, Null, Nothing)"
    TestAssert ParseCall(strExpr, arrParams), "mixed parameters are accepted"
    TestAssert strExpr = "MyProc", "name extracted"
    TestAssert UBound(arrParams) = 4, "five parameters"
    TestAssert arrParams(0) = "a, b", "comma inside double quotes"
    TestAssert arrParams(1) = "c", "single quotes removed"
    TestAssert arrParams(2) = "12", "unquoted value passed as string"
    TestAssert IsNull(arrParams(3)), "Null literal"
    TestAssert arrParams(4) Is Nothing, "Nothing literal"

    strExpr = " MyProc ( ""x"" ) "
    TestAssert ParseCall(strExpr, arrParams), "surrounding whitespace is accepted"
    TestAssert strExpr = "MyProc", "name trimmed"
    TestAssert arrParams(0) = "x", "single parameter"

    strExpr = "MyProc(""ab'c"", 'x')"
    TestAssert ParseCall(strExpr, arrParams), "single quote inside double quotes"
    TestAssert UBound(arrParams) = 1, "two parameters"
    TestAssert arrParams(0) = "ab'c", "single quote kept in string"
    TestAssert arrParams(1) = "x", "following single-quoted parameter"

    strExpr = "MyProc(""ab''c"", ""ab""""c"", 'ab''c')"
    TestAssert ParseCall(strExpr, arrParams), "doubled quotes inside strings"
    TestAssert UBound(arrParams) = 2, "three parameters with doubled quotes"
    TestAssert arrParams(0) = "ab''c", "doubled single quote kept inside double quotes"
    TestAssert arrParams(1) = "ab""c", "doubled double quote becomes one"
    TestAssert arrParams(2) = "ab'c", "doubled single quote becomes one inside single quotes"

    strExpr = "MyProc("""")"
    TestAssert ParseCall(strExpr, arrParams), "empty string parameter"
    TestAssert arrParams(0) = vbNullString, "empty string value"

End Sub


Public Sub TestProcedureCallSyntaxErrors()

    Dim arrParams() As Variant

    TestAssert Not TryParse("MyProc(""a""", arrParams), "missing closing parenthesis"
    TestAssert Not TryParse("MyProc(a,,b)", arrParams), "empty value between commas"
    TestAssert Not TryParse("MyProc(a,)", arrParams), "trailing comma"
    TestAssert Not TryParse("MyProc(""a)", arrParams), "unclosed quote"
    TestAssert Not TryParse("MyProc(""a"" x)", arrParams), "text after quoted string"
    TestAssert Not TryParse("MyProc(""ab""c"")", arrParams), "single quote char inside string not doubled"
    TestAssert Not TryParse("MyProc(a) x", arrParams), "text after closing parenthesis"
    TestAssert Not TryParse("(a)", arrParams), "missing procedure name"

End Sub


Public Sub TestSplitProcedureList()

    Dim colList As Collection

    Set colList = SplitList("ProcA")
    TestAssert colList.Count = 1, "single procedure"
    TestAssert colList(1) = "ProcA", "single procedure name"

    Set colList = SplitList(" ProcA : ProcB ")
    TestAssert colList.Count = 2, "two procedures"
    TestAssert colList(1) = "ProcA", "first entry trimmed"
    TestAssert colList(2) = "ProcB", "second entry trimmed"

    Set colList = SplitList("ProcA(""x:y"") : ProcB('a:b')")
    TestAssert colList.Count = 2, "colon inside quotes does not separate"
    TestAssert colList(1) = "ProcA(""x:y"")", "double-quoted colon kept"
    TestAssert colList(2) = "ProcB('a:b')", "single-quoted colon kept"

    Set colList = SplitList("ProcA(""ab""""c:d"")")
    TestAssert colList.Count = 1, "doubled quote does not end the string"

    Set colList = SplitList("ProcA(a:b)")
    TestAssert colList.Count = 1, "colon inside parentheses does not separate"

    Set colList = SplitList("C:\Path\AddIn.Proc : D:/Path/AddIn.Proc2")
    TestAssert colList.Count = 2, "drive colons do not separate"
    TestAssert colList(1) = "C:\Path\AddIn.Proc", "first path kept"
    TestAssert colList(2) = "D:/Path/AddIn.Proc2", "second path kept"

    Set colList = SplitList("ProcA:\\Server\Share\AddIn.Proc")
    TestAssert colList.Count = 2, "colon before UNC path separates"
    TestAssert colList(2) = "\\Server\Share\AddIn.Proc", "UNC path kept"

    Set colList = SplitList("ProcA::ProcB:")
    TestAssert colList.Count = 2, "empty entries skipped"

    Set colList = SplitList(vbNullString)
    TestAssert colList.Count = 0, "empty option value"

End Sub


Public Sub TestFindAddInCall()

    Dim strRoot As String
    Dim strFolder As String
    Dim strAddInPath As String
    Dim strCall As String

    ' Folder names with a dot and parentheses must not end the add-in path
    strRoot = GetTempFolder("VCS")
    strFolder = BuildPath2(strRoot, "Folder (x86)")
    FSO.CreateFolder strFolder
    strFolder = BuildPath2(strFolder, "My.AddIns")
    FSO.CreateFolder strFolder
    ' Only the file's existence is checked (WriteFile would not create an empty file)
    WriteFile "dummy", BuildPath2(strFolder, "TestAddIn.accda")

    With New clsProcedureRunner
        TestAssert .FindAddInCall("ProcA", strAddInPath, strCall) = eacNone, "plain name is no add-in call"
        TestAssert .FindAddInCall("Module1.ProcA", strAddInPath, strCall) = eacNone, "module prefix is no add-in call"

        TestAssert .FindAddInCall(BuildPath2(strFolder, "TestAddIn.HookProc(""a.b"", 1)"), strAddInPath, strCall) = eacFound, _
            "add-in found"
        TestAssert strAddInPath = BuildPath2(strFolder, "TestAddIn"), "add-in path without extension"
        TestAssert strCall = "HookProc(""a.b"", 1)", "call expression with parameters"

        TestAssert .FindAddInCall(BuildPath2(strFolder, "TestAddIn.HookProc"), strAddInPath, strCall) = eacFound, _
            "add-in found without parameters"
        TestAssert strCall = "HookProc", "call expression without parameters"

        TestAssert .FindAddInCall(BuildPath2(strFolder, "Missing.HookProc"), strAddInPath, strCall) = eacFileMissing, _
            "missing add-in file"

        WriteFile "dummy", BuildPath2(strFolder, "Compiled.accde")
        TestAssert .FindAddInCall(BuildPath2(strFolder, "Compiled.HookProc"), strAddInPath, strCall) = eacFound, _
            "compiled add-in (.accde) found"
        TestAssert strAddInPath = BuildPath2(strFolder, "Compiled"), "compiled add-in path without extension"

        WriteFile "dummy", BuildPath2(strFolder, "TestAddIn.accde")
        TestAssert .FindAddInCall(BuildPath2(strFolder, "TestAddIn.HookProc"), strAddInPath, strCall) = eacAmbiguous, _
            ".accda and .accde with the same name"

        TestAssert .ExpandAddInPath("%addins%\X") = BuildPath2(Environ$("AppData"), "Microsoft", "AddIns", "X"), _
            "%addins% placeholder"
        TestAssert .ExpandAddInPath("%APPDATA%\X") = Environ$("AppData") & "\X", "environment variable"
    End With

    FSO.DeleteFolder strRoot, True

End Sub


Public Sub TestProcedureContextWithoutListener()

    Dim oContext As clsVCSProcedureContext
    Dim arrParams() As Variant

    With New clsProcedureRunner
        Set oContext = .NewProcedureContext("TestHook")
        ' Outside RunExternalProcedure the VCS keyword passes Nothing
        TestAssert .TryGetProcedureNameAndParams("MyProc(VCS)", arrParams), "VCS keyword accepted"
        TestAssert arrParams(0) Is Nothing, "VCS is Nothing outside RunExternalProcedure"
    End With

    TestAssert oContext.HookName = "TestHook", "hook name"
    TestAssert oContext.DatabasePath = CurrentProject.FullName, "database path"
    TestAssert Len(oContext.AddInVersion) > 0, "add-in version"

    ' No listener: writes do nothing, live values return defaults
    oContext.AddLog "not logged"
    oContext.LogError eelError, "not logged"
    TestAssert oContext.ErrorLevel = eelNoError, "error level default without listener"
    TestAssert oContext.InteractionMode = eimNormal, "interaction mode default without listener"

End Sub


Public Sub TestProcedureContextListener()

    Dim oContext As clsVCSProcedureContext
    Dim oListener As clsVCSProcedureContextListener
    Dim eelSavedLevel As eErrorLevel
    Dim lngJournalStart As Long
    Dim colEntries As Collection

    With New clsProcedureRunner
        Set oContext = .NewProcedureContext("TestHook")
    End With
    Set oListener = New clsVCSProcedureContextListener
    oListener.Attach oContext

    ' Logged levels raise the operation's error level; restore it afterwards so the
    ' test runner does not count this test as errored.
    eelSavedLevel = Operation.ErrorLevel
    lngJournalStart = Log.ErrorJournalCount

    TestAssert oContext.ErrorLevel = Operation.ErrorLevel, "error level from operation"
    TestAssert oContext.InteractionMode = Operation.InteractionMode, "interaction mode from operation"

    oContext.LogError eelWarning, "context warning", "TestSource"
    oContext.LogError 99, "unknown level"
    Set colEntries = Log.GetErrorJournalSince(lngJournalStart)
    TestAssert colEntries.Count = 2, "two journal entries"
    TestAssert colEntries(1)("level") = "warning", "level passed through"
    TestAssert colEntries(1)("source") = "TestSource", "source passed through"
    TestAssert colEntries(2)("level") = "alert", "unknown level logged as alert"
    TestAssert colEntries(2)("source") = "TestHook", "missing source replaced by hook name"

    ' After Detach nothing reaches the log any more
    oListener.Detach
    lngJournalStart = Log.ErrorJournalCount
    oContext.LogError eelWarning, "after detach"
    TestAssert Log.ErrorJournalCount = lngJournalStart, "nothing logged after detach"
    TestAssert oContext.ErrorLevel = eelNoError, "error level default after detach"

    Operation.ErrorLevel = eelSavedLevel

End Sub


Public Sub TestParseReturnMessage()

    Dim strMessage As String

    With New clsProcedureRunner
        TestAssert .ParseReturnMessage("Error: failed to copy", strMessage) = eelError, "Error prefix"
        TestAssert strMessage = "failed to copy", "prefix removed"
        TestAssert .ParseReturnMessage("warning: x", strMessage) = eelAlert, "Warning prefix (case-insensitive)"
        TestAssert .ParseReturnMessage("Alert: x", strMessage) = eelAlert, "Alert prefix"
        TestAssert .ParseReturnMessage("Failed: x", strMessage) = eelAlert, "Failed prefix"
        TestAssert .ParseReturnMessage("Critical: x", strMessage) = eelCritical, "Critical prefix"
        TestAssert .ParseReturnMessage("FATAL: x", strMessage) = eelCritical, "Fatal prefix"
        TestAssert .ParseReturnMessage("Info: x", strMessage) = eelInfo, "Info prefix"
        TestAssert .ParseReturnMessage("Note: x", strMessage) = eelInfo, "Note prefix"
        TestAssert .ParseReturnMessage("Success: x", strMessage) = eelInfo, "Success prefix"
        TestAssert .ParseReturnMessage("Log: x", strMessage) = eelWarning, "Log prefix (log file only)"
        TestAssert strMessage = "x", "message after prefix"

        TestAssert .ParseReturnMessage("Copied 3 files", strMessage) = eelInfo, "text without prefix is info"
        TestAssert strMessage = "Copied 3 files", "full text kept"
        TestAssert .ParseReturnMessage("Path: C:\Temp", strMessage) = eelInfo, "unknown prefix is info"
        TestAssert strMessage = "Path: C:\Temp", "unknown prefix kept"
        TestAssert .ParseReturnMessage(":x", strMessage) = eelInfo, "leading colon is no prefix"
        TestAssert .ParseReturnMessage("   ", strMessage) = eelNoError, "empty text logs nothing"
    End With

End Sub


Public Sub TestScalarOrEmpty()
    With New clsProcedureRunner
        TestAssert .ScalarOrEmpty(5) = 5, "number passed through"
        TestAssert .ScalarOrEmpty("x") = "x", "text passed through"
        TestAssert .ScalarOrEmpty(False) = False, "boolean passed through"
        TestAssert IsEmpty(.ScalarOrEmpty(New Collection)), "object returns Empty"
        TestAssert IsEmpty(.ScalarOrEmpty(Nothing)), "Nothing returns Empty"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SplitList
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Split a Run* option value as RunExternalProcedure does.
'---------------------------------------------------------------------------------------
'
Private Function SplitList(ByVal strOptionValue As String) As Collection
    With New clsProcedureRunner
        Set SplitList = .SplitProcedureList(strOptionValue)
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseCall
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Full preparation as used by RunProcInCurrentProject. Use only with valid
'           : expressions; an invalid one logs an error and shows a message box.
'---------------------------------------------------------------------------------------
'
Private Function ParseCall(ByRef strExpr As String, ByRef arrParams() As Variant) As Boolean
    Erase arrParams
    With New clsProcedureRunner
        ParseCall = .PrepareProcedureCallExpressionAndParams(strExpr, arrParams)
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : TryParse
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Parser only, without logging.
'---------------------------------------------------------------------------------------
'
Private Function TryParse(ByVal strExpr As String, ByRef arrParams() As Variant) As Boolean
    Erase arrParams
    With New clsProcedureRunner
        TryParse = .TryGetProcedureNameAndParams(strExpr, arrParams)
    End With
End Function

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

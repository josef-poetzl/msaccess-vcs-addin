VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "modImportExportTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
' Runs before each test
End Sub

Public Sub TearDown()
' Runs after each test
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("%addins%\AddInName1.Proc1:%addins%\AddInName2.Proc2", "%addins%\AddInName1.Proc1|%addins%\AddInName2.Proc2")
'AccUnit:Row("%addins%\AddInName1.Proc1 : %addins%\AddInName2.Proc2", "%addins%\AddInName1.Proc1|%addins%\AddInName2.Proc2")
'AccUnit:Row("Proc1:Proc2", "Proc1|Proc2")
'AccUnit:Row("Proc1(""a:B"")", "Proc1(""a:B"")")
'AccUnit:Row("Proc1(""""a:B"""")", "Proc1(""""a:B"""")")
'AccUnit:Row("Proc1", "Proc1")
'AccUnit:Row("%addins%\AddInName1.Proc1", "%addins%\AddInName1.Proc1")
'AccUnit:Row(":", "")
'AccUnit:Row("   :   :   ", "")
'AccUnit:Row("   : Proc1  :   ", "Proc1")
Public Sub GetProcedureCalls_RexExPatternCheck(ByVal RunProceduresString As String, ByRef Expected As String)
    
    Const RegExPattern As String = "[^:()]+(?:\([^)]*\))?" ' split by : + ignore : inside " e.g. proc("a:b"):proc2
    
    Dim Actual As String
    Dim AcutalArray() As String
    AcutalArray = clsRegExpSupport.RegExSplit(RunProceduresString, RegExPattern)
    Actual = Join(AcutalArray, "|")
            
    Assert.That Actual, Iz.EqualTo(Expected)
    
End Sub




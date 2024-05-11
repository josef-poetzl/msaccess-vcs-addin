VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVersionControlRunBeforeExportTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_VCS As Object
Private m_TestBridge As TestHelperCallbackBridge

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
    Set m_VCS = GetVCS
    m_VCS.SilentMode = True
    Set m_TestBridge = New TestHelperCallbackBridge
    modTestHelper_RunBeforeTestProc.SetTestBridge m_TestBridge
End Sub

Public Sub TearDown()

    If Err.Number <> 0 Then 'remove error: VCS checks err object!
       Err.Clear
    End If
    
    modTestHelper_RunBeforeTestProc.SetTestBridge Nothing
    Set m_TestBridge = Nothing
    
    m_VCS.Options.RunBeforeExport = vbNullString
    m_VCS.Options.SaveOptionsForProject
    
    Set m_VCS = Nothing
    
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub TestRunBeforeExport_CallSub_CheckCall()

    Const ProcedureToStart As String = "RunBeforeExportTestFunction1"
    Dim VcsErrorLevel As Long
    
    m_VCS.Options.RunBeforeExport = ProcedureToStart
    m_VCS.Options.SaveOptionsForProject
    VcsErrorLevel = m_VCS.ExportVBA
    
    Assert.That m_TestBridge.CalledProcedure, Iz.EqualTo(ProcedureToStart)

End Sub

Public Sub TestRunBeforeExport_CallFunction_CheckCallAndErrorLevel()

    Const ProcedureToStart As String = "RunBeforeExportTestFunction2"
    Dim VcsErrorLevel As Long
    
    m_TestBridge.FunctionReturnValue = True
    
    m_VCS.Options.RunBeforeExport = ProcedureToStart
    m_VCS.Options.SaveOptionsForProject
    VcsErrorLevel = m_VCS.ExportVBA
    
    Assert.That m_TestBridge.CalledProcedure, Iz.EqualTo(ProcedureToStart)

End Sub

'AccUnit:Row(True, 0).Name = "True"
'AccUnit:Row(False, 5).Name = "False"
Public Sub TestRunBeforeExport_CallBooleanFunction_CheckCallAndErrorLevelFailed(ByVal FunctionReturnValue As Boolean, ByVal ExpectedErrorLevel As Long)

    Const ProcedureToStart As String = "RunBeforeExportTestFunction2"
    Dim VcsErrorLevel As Long
    
    m_TestBridge.FunctionReturnValue = FunctionReturnValue
    
    m_VCS.Options.RunBeforeExport = ProcedureToStart
    m_VCS.Options.SaveOptionsForProject
    VcsErrorLevel = m_VCS.ExportVBA
    
    Assert.That m_TestBridge.CalledProcedure, Iz.EqualTo(ProcedureToStart)
    Assert.That VcsErrorLevel, Iz.EqualTo(ExpectedErrorLevel)

End Sub

'AccUnit:Row("Log: ...", 1, False).Name = "Log"
'AccUnit:Row("Info: ...", 2, False).Name = "Info"
'AccUnit:Row("Success: ...", 2, False).Name = "Success"
'AccUnit:Row("Note: ...", 2, False).Name = "Note"
'AccUnit:Row("Warning: ...", 3, True).Name = "Warning"
'AccUnit:Row("Alert: ...", 3, True).Name = "Alert"
'AccUnit:Row("Failed: ...", 3, True).Name = "Failed"
'AccUnit:Row("Error: ...", 4, true).Name = "Error"
'AccUnit:Row("Critical: ...", 5, true).Name = "Critical"
Public Sub TestRunBeforeExport_CallStringFunction_CheckCallAndErrorLevelFailed( _
                            ByVal FunctionReturnValue As String, _
                            ByVal ExpectedErrorLevel As Long, _
                            ByVal ExpectedVcsMainFormIsOpen As Boolean)

    Const ProcedureToStart As String = "RunBeforeExportTestFunction3"
    Dim VcsErrorLevel As Long
    
    m_TestBridge.FunctionReturnValue = FunctionReturnValue
    
    m_VCS.Options.RunBeforeExport = ProcedureToStart
    m_VCS.Options.SaveOptionsForProject
    VcsErrorLevel = m_VCS.ExportVBA

    Assert.That m_TestBridge.CalledProcedure, Iz.EqualTo(ProcedureToStart)
    Assert.That VcsErrorLevel, Iz.EqualTo(ExpectedErrorLevel)
    Assert.That VcsMainFormIsOpen, Iz.EqualTo(ExpectedVcsMainFormIsOpen)

End Sub

'--------------------------------------------------------------------
' Test helper procedures
'--------------------------------------------------------------------
Private Function GetVCS() As Object
   Set GetVCS = RunOjectAddInProcedure("VCS")
End Function

Private Function GetLog() As Object
   Set GetLog = RunOjectAddInProcedure("Log")
End Function

Private Function RunOjectAddInProcedure(ByVal ProcedureName As String) As Object
   Set RunOjectAddInProcedure = Application.Run(GetMsAccessVcsPath & "\Version Control." & ProcedureName)
End Function

Private Function GetMsAccessVcsPath() As String
   GetMsAccessVcsPath = Environ("Appdata") & "\MsAccessVCS"
End Function

Private Function VcsMainFormIsOpen() As Boolean
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmVCSMain" Then
            VcsMainFormIsOpen = True
            Exit Function
        End If
    Next
    
End Function



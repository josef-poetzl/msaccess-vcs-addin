Attribute VB_Name = "modTestHelper_RunBeforeTestProc"
Option Compare Database
Option Explicit

' AccUnit:TestRelated

Private m_TestBridge As TestHelperCallbackBridge

Public Sub SetTestBridge(ByVal NewRef As TestHelperCallbackBridge)
   Set m_TestBridge = NewRef
End Sub

Public Sub RunBeforeExportTestFunction1()
   If m_TestBridge Is Nothing Then
      Debug.Print "RunBeforeExportTestFunction1 called"
   Else
      m_TestBridge.CommitProcedureCall "RunBeforeExportTestFunction1"
   End If
End Sub

Public Function RunBeforeExportTestFunction2() As Boolean
   If m_TestBridge Is Nothing Then
      Debug.Print "RunBeforeExportTestFunction2 called"
   Else
      m_TestBridge.CommitProcedureCall "RunBeforeExportTestFunction2"
      RunBeforeExportTestFunction2 = CBool(m_TestBridge.FunctionReturnValue)
   End If
End Function

Public Function RunBeforeExportTestFunction3() As String
   If m_TestBridge Is Nothing Then
      Debug.Print "RunBeforeExportTestFunction3 called"
   Else
      m_TestBridge.CommitProcedureCall "RunBeforeExportTestFunction3"
      RunBeforeExportTestFunction3 = m_TestBridge.FunctionReturnValue
   End If
End Function

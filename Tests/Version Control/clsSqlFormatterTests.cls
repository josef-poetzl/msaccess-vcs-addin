VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSqlFormatterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass
' Call: TestSuite.AddByClassName("clsSqlFormatterTests").Run

Private m_SqlFormatter As clsSqlFormatter

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
   Set m_SqlFormatter = New clsSqlFormatter
End Sub

Public Sub TearDown()
   Set m_SqlFormatter = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub TestSqlString_SingleQuote()

   Const UnformattedSql As String = "Select 'xxx\' & FileName As FullFileName, 'a''bc' as T1, ""a'bc"" as T2 From FileTable"
   Const Expected As String = _
                  "Select" & vbNewLine & _
                  "  'xxx\' & FileName As FullFileName," & vbNewLine & _
                  "  'a''bc' as T1," & vbNewLine & _
                  "  ""a'bc"" as T2" & vbNewLine & _
                  "From" & vbNewLine & _
                  "  FileTable"
   
   Dim Actual As String
   Actual = m_SqlFormatter.FormatSQL(UnformattedSql)

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub TestSqlString_DoubleQuote()

   Const UnformattedSql As String = "Select ""xxx\"" & FileName As FullFileName, ""a""""bc"" as T1, ""a'bc"" as T2 From FileTable"
   Const Expected As String = _
                  "Select" & vbNewLine & _
                  "  ""xxx\"" & FileName As FullFileName," & vbNewLine & _
                  "  ""a""""bc"" as T1," & vbNewLine & _
                  "  ""a'bc"" as T2" & vbNewLine & _
                  "From" & vbNewLine & _
                  "  FileTable"
   
   Dim Actual As String
   Actual = m_SqlFormatter.FormatSQL(UnformattedSql)

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub TestSqlString_Like()

   Const UnformattedSql As String = "Select * From Table Where T like 'a[1-5]z'"
   Const Expected As String = _
                  "Select" & vbNewLine & _
                  "  *" & vbNewLine & _
                  "From" & vbNewLine & _
                  "  Table" & vbNewLine & _
                  "Where" & vbNewLine & _
                  "  T like 'a[1-5]z'"
   
   Dim Actual As String
   Actual = m_SqlFormatter.FormatSQL(UnformattedSql)

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub


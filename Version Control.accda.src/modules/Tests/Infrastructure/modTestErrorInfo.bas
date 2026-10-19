Attribute VB_Name = "modTestErrorInfo"
'---------------------------------------------------------------------------------------
' Module    : modTestErrorInfo
' Author    : Josef Pötzl
' Date      : 10/19/2026
' Purpose   : Tests for clsErrorInfo: display data per error level and the order
'           : of eErrorLevel. (Nothing is logged, so the operation's error level
'           : is not changed.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestErrorLevelOrder()
    TestAssert eelNoError < eelWarning, "warning above no error"
    TestAssert eelWarning < eelInfo, "info above warning"
    TestAssert eelInfo < eelAlert, "alert above info"
    TestAssert eelAlert < eelError, "error above alert (info/alert still count as success)"
    TestAssert eelError < eelCritical, "critical above error"
End Sub


Public Sub TestErrorInfoDisplayData()

    With NewErrorInfo(eelInfo, "text")
        TestAssert .DisplayText = "INFO: text", "info prefix"
        TestAssert .ErrorLevelText = "Info", "info level text"
        TestAssert .LogColor = "blue", "info color"
        TestAssert .MsgBoxStyle = vbInformation, "info message box style"
    End With

    With NewErrorInfo(eelAlert, "text")
        TestAssert .DisplayText = "WARNING: text", "alert prefix"
        TestAssert .ErrorLevelText = "Warning", "alert level text"
        TestAssert .LogColor = "#CC7700", "alert color"
        TestAssert .MsgBoxStyle = vbInformation, "alert message box style"
    End With

    With NewErrorInfo(eelWarning, "text")
        TestAssert .DisplayText = "WARNING: text", "warning prefix"
        TestAssert .LogColor = "red", "warning color unchanged"
    End With

    With NewErrorInfo(eelError, "text")
        TestAssert .DisplayText = "ERROR: text", "error prefix"
        TestAssert .LogColor = "red", "error color"
        TestAssert .MsgBoxStyle = vbExclamation, "error message box style"
    End With

    With NewErrorInfo(eelCritical, "text")
        TestAssert .DisplayText = "CRITICAL: text", "critical prefix"
        TestAssert .LogColor = "red", "critical color"
    End With

End Sub

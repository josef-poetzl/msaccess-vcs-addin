VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegExpSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Friend Function RegExSplit(ByVal strTextToSplit As String, ByVal strPattern As String) As String()

    Dim arrResult() As String
    Dim Matches As MatchCollection
    Dim RegexMatch As Match
    Dim strProcedureCall As String
    Dim i As Long
    Dim lngProcCallCnt As Long
    Dim lngMatches As Long

    With New RegExp
        .IgnoreCase = True
        .Global = True
        .Pattern = strPattern

        If Not .Test(strTextToSplit) Then
            ReDim arrResult(0)
            RegExSplit = arrResult
            Exit Function
        End If

        With .Execute(strTextToSplit)
            lngMatches = .Count
            ReDim arrResult(lngMatches - 1)
            For i = 0 To .Count - 1
                strProcedureCall = Trim(.Item(i))
                If Len(strProcedureCall) > 0 Then
                    arrResult(lngProcCallCnt) = strProcedureCall
                    lngProcCallCnt = lngProcCallCnt + 1
                End If
            Next
        End With

        If lngProcCallCnt < lngMatches Then
            If lngProcCallCnt = 0 Then
                ReDim arrResult(0)
            Else
                ReDim Preserve arrResult(lngProcCallCnt - 1)
            End If
        End If

    End With

    RegExSplit = arrResult

End Function

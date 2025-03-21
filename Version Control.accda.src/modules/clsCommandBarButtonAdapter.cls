VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandBarButtonAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents m_CommandBarButton As CommandBarButton
Attribute m_CommandBarButton.VB_VarHelpID = -1
Private m_EventBridge As clsCommandBarAdapterEventBridge
Private m_CommandCode As String

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_EventBridge = Nothing
    m_CommandBarButton.Delete
    Set m_CommandBarButton = Nothing
End Sub

Public Function Init(ByVal objEventBridge As clsCommandBarAdapterEventBridge, ByVal strCommandCode As String, _
                     ByVal objCommandBars As CommandBars, ByVal strCommandBarName As String, _
                     ByVal strCaption As String, ByVal FaceId As Long, ByVal lngBeforeId As Long, _
            Optional ByVal bolBeginGroup As Boolean = False)

    Set m_EventBridge = objEventBridge
    m_CommandCode = strCommandCode
    InitCommandBarButton objCommandBars, strCommandBarName, strCaption, FaceId, lngBeforeId, bolBeginGroup

End Function

Public Property Get Self() As clsCommandBarButtonAdapter
    Set Self = Me
End Property

Public Property Get CommandCode() As String
    CommandCode = m_CommandCode
End Property

Public Property Get CommandBarButton() As CommandBarButton
    Set CommandBarButton = m_CommandBarButton
End Property

Private Function InitCommandBarButton(ByVal objCommandBars As CommandBars, ByVal strCommandBarName As String, _
                                      ByVal strCaption As String, ByVal FaceId As Long, ByVal lngBeforeId As Long, _
                             Optional ByVal bolBeginGroup As Boolean = False)

    Dim cb As CommandBar
    Dim cbb As CommandBarButton
    Dim lngBeforeIndex As Long
    Dim strTag As String

    strTag = "MsAccesVCS." & strCommandBarName & "." & strCaption

    Set cb = objCommandBars(strCommandBarName)
    Set cbb = FindCommandbarControlByTag(cb, strTag)

    If Not (cbb Is Nothing) Then
        Set m_CommandBarButton = cbb
        Exit Function
    End If

    If lngBeforeId > 0 Then
        lngBeforeIndex = GetButtonIndex(cb, lngBeforeId)
        If lngBeforeIndex = -1 Then
            lngBeforeIndex = 0
        End If
    End If

    If lngBeforeIndex > 0 Then
        Set cbb = cb.Controls.Add(MsoControlType.msoControlButton, , , lngBeforeIndex, True)
    Else
        Set cbb = cb.Controls.Add(MsoControlType.msoControlButton, , , , True)
    End If
    cbb.Caption = strCaption
    cbb.BeginGroup = bolBeginGroup
    cbb.Tag = strTag
    cbb.FaceId = FaceId

    Set m_CommandBarButton = cbb

End Function

Private Function FindCommandbarControlByTag(ByVal cmdBar As CommandBar, ByVal strTag As String) As CommandBarControl

    Set FindCommandbarControlByTag = cmdBar.FindControl(MsoControlType.msoControlButton, , strTag)

End Function

Private Function GetButtonIndex(ByVal cmdBar As CommandBar, ByVal ControlID As Long)

    Dim cbctl As CommandBarControl
    Set cbctl = cmdBar.FindControl(MsoControlType.msoControlButton, ControlID)
    If cbctl Is Nothing Then
        GetButtonIndex = -1
    Else
        GetButtonIndex = cbctl.Index
    End If

End Function

Private Sub m_CommandBarButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    m_EventBridge.RaiseClicked Me
End Sub

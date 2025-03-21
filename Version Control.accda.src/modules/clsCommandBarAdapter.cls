VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandBarAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_ButtonAdapters As Collection
Private WithEvents m_EventBridge As clsCommandBarAdapterEventBridge
Attribute m_EventBridge.VB_VarHelpID = -1

Private Const CompareSelectedWithBaseCommandCode As String = "CompareSelectedWithBase"
Private Const ExportSelectedCommandCode As String = "ExportSelected"
Private Const LoadSelectedCommandCode As String = "LoadSelected"

Private Sub Class_Initialize()
    Set m_ButtonAdapters = New Collection
    Set m_EventBridge = New clsCommandBarAdapterEventBridge
    AddCommandBarButtons
End Sub

Private Sub Class_Terminate()
    Set m_EventBridge = Nothing
    Set m_ButtonAdapters = Nothing
End Sub


Private Sub AddCommandBarButtons()

'FaceIds Compare: 25, 202, 720, 2652

' Compare with base
    AddSelectedCommandButtons CompareSelectedWithBaseCommandCode, "Compare with base", 2652, True

' Export / Load
    AddSelectedCommandButtons ExportSelectedCommandCode, "VCS Export", 2116
    AddSelectedCommandButtons LoadSelectedCommandCode, "VCS Load", 2109


End Sub

Private Sub AddSelectedCommandButtons(ByVal strCommandCode As String, ByVal strCaption As String, ByVal FaceId As Long, Optional ByVal bolBeginGroup As Boolean = False)

    AddCommandBarButton strCommandCode, Application.CommandBars, "Navigation Pane View Pop-up", strCaption, FaceId, 0, bolBeginGroup
    AddCommandBarButton strCommandCode, Application.CommandBars, "Navigation Pane object Pop-up", strCaption, FaceId, 0, bolBeginGroup

    AddCommandBarButton strCommandCode, VBE.CommandBars, "Code Window", strCaption, FaceId, 473, bolBeginGroup
    AddCommandBarButton strCommandCode, VBE.CommandBars, "Project Window", strCaption, FaceId, 2578, bolBeginGroup

End Sub

Private Sub AddCommandBarButton(ByVal strCommandCode As String, ByVal objCommandBars As CommandBars, ByVal strCommandBarName As String, _
                                ByVal strCaption As String, ByVal FaceId As Long, ByVal lngBeforeId As Long, _
                       Optional ByVal bolBeginGroup As Boolean = False)

    With New clsCommandBarButtonAdapter
        .Init m_EventBridge, strCommandCode, objCommandBars, strCommandBarName, strCaption, FaceId, lngBeforeId, bolBeginGroup
        m_ButtonAdapters.Add .Self
    End With

End Sub

Private Sub m_EventBridge_Clicked(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter)
    RunCommandFromButtonAdapter objCommandBarButtonAdapter
End Sub

Private Sub RunCommandFromButtonAdapter(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter)

    Select Case objCommandBarButtonAdapter.CommandCode
        Case CompareSelectedWithBaseCommandCode, ExportSelectedCommandCode, LoadSelectedCommandCode
            RunSelectedCommand GetSelectedObjectFromButtonAdapter(objCommandBarButtonAdapter), objCommandBarButtonAdapter.CommandCode
    End Select

End Sub

Private Function GetSelectedObjectFromButtonAdapter(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter) As AccessObject

    If objCommandBarButtonAdapter.CommandBarButton.Parent.Parent Is Access.Application Then
        Set GetSelectedObjectFromButtonAdapter = GetSelectedObjectFromAccessApplication
    Else ' VBE
        Set GetSelectedObjectFromButtonAdapter = GetSelectedObjectFromVBE
    End If

End Function

Private Function GetSelectedObjectFromAccessApplication() As AccessObject

    Dim objSelectedObject As AccessObject

    Select Case Application.CurrentObjectType
        Case AcObjectType.acModule
            Set objSelectedObject = CurrentProject.AllModules(Application.CurrentObjectName)
        Case AcObjectType.acForm
            Set objSelectedObject = CurrentProject.AllForms(Application.CurrentObjectName)
        Case AcObjectType.acReport
            Set objSelectedObject = CurrentProject.AllReports(Application.CurrentObjectName)
        Case AcObjectType.acMacro
            Set objSelectedObject = CurrentProject.AllMacros(Application.CurrentObjectName)
        Case Else
            Exit Function

    End Select

    Set GetSelectedObjectFromAccessApplication = objSelectedObject

End Function

Private Function GetSelectedObjectFromVBE() As AccessObject

    Dim vbc As VBComponent
    Dim objSelectedObject As AccessObject

    Set vbc = VBE.SelectedVBComponent

    If vbc.Type = vbext_ct_Document Then
        If Left(vbc.Name, 5) = "Form_" Then
            Set objSelectedObject = CurrentProject.AllForms(Mid(vbc.Name, Len("Form_") + 1))
        ElseIf Left(vbc.Name, 7) = "Report_" Then
            Set objSelectedObject = CurrentProject.AllReports(Mid(vbc.Name, Len("Report_") + 1))
        End If
    Else
        Set objSelectedObject = CurrentProject.AllModules(vbc.Name)
    End If

    Set GetSelectedObjectFromVBE = objSelectedObject

End Function


Private Sub RunSelectedCommand(ByVal objSelected As AccessObject, ByVal strCommandCode As String)

    Select Case strCommandCode
        Case CompareSelectedWithBaseCommandCode
            CompareObjectWithBase objSelected

        Case ExportSelectedCommandCode
            ExportObject objSelected

        Case LoadSelectedCommandCode
            LoadObject objSelected

    End Select

End Sub

Private Sub CompareObjectWithBase(ByVal objToCompare As AccessObject)

    Dim strSourceFilePath As String

    If objToCompare Is Nothing Then
        Exit Sub
    End If

    modImportExport.ExportSingleObject objToCompare

    With GetClassFromObject(objToCompare)
        strSourceFilePath = .SourceFile
        If .ComponentType = edbForm Then
            strSourceFilePath = Left(strSourceFilePath, Len(strSourceFilePath) - Len(".bas")) & ".cls"
        End If
    End With

    modObjects.Diff.Files strSourceFilePath, vbNullString

End Sub

Private Sub ExportObject(ByVal objToExport As AccessObject)
    VCS.ExportObject objToExport
End Sub

Private Sub LoadObject(ByVal objToLoad As AccessObject)
    VCS.LoadObject objToLoad
End Sub

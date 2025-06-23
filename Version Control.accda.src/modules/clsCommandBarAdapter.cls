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
        Case CompareSelectedWithBaseCommandCode
            RunSelectedCommand GetSelectedObjectsFromButtonAdapter(objCommandBarButtonAdapter, True), objCommandBarButtonAdapter.CommandCode
        Case ExportSelectedCommandCode, LoadSelectedCommandCode
            RunSelectedCommand GetSelectedObjectsFromButtonAdapter(objCommandBarButtonAdapter), objCommandBarButtonAdapter.CommandCode
    End Select

End Sub

Private Function GetSelectedObjectsFromButtonAdapter(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter, _
                                            Optional ByVal DisableObjectsFromVbeProjectTreeGroup As Boolean = False) As AccessObject()

    If objCommandBarButtonAdapter.CommandBarButton.Parent.Parent Is Access.Application Then
        GetSelectedObjectsFromButtonAdapter = GetSelectedObjectFromAccessApplication
    Else ' VBE
        GetSelectedObjectsFromButtonAdapter = GetSelectedObjectsFromVBE
    End If

End Function

Private Function GetSelectedObjectFromAccessApplication() As AccessObject()

    Dim objSelectedObject(0) As AccessObject

    Select Case Application.CurrentObjectType
        Case AcObjectType.acModule
            Set objSelectedObject(0) = CurrentProject.AllModules(Application.CurrentObjectName)
        Case AcObjectType.acForm
            Set objSelectedObject(0) = CurrentProject.AllForms(Application.CurrentObjectName)
        Case AcObjectType.acReport
            Set objSelectedObject(0) = CurrentProject.AllReports(Application.CurrentObjectName)
        Case AcObjectType.acMacro
            Set objSelectedObject(0) = CurrentProject.AllMacros(Application.CurrentObjectName)
        Case Else
            Exit Function

    End Select

    GetSelectedObjectFromAccessApplication = objSelectedObject

End Function

Private Function GetSelectedObjectsFromVBE(Optional ByVal DisableObjectsFromVbeProjectTreeGroup As Boolean = False) As AccessObject()

    Dim SelectedVbc As VBComponent
    Dim SelectedTreeItemLevel As Long
    Dim arrVbc() As VBComponent
    Dim i As Long
    Dim objSelectedObject() As AccessObject

    Set SelectedVbc = VBE.SelectedVBComponent

    If SelectedVbc Is Nothing Then
        If DisableObjectsFromVbeProjectTreeGroup Then
            Exit Function
        End If
        arrVbc = GetSelectedVBComponentsFromVbeProjectTreeGroup()
    Else
        ReDim arrVbc(0)
        Set arrVbc(0) = SelectedVbc
    End If

    ReDim objSelectedObject(UBound(arrVbc))

    For i = LBound(arrVbc) To UBound(arrVbc)
        Set objSelectedObject(i) = GetAccessObjectOfVbComponente(arrVbc(i))
    Next

    GetSelectedObjectsFromVBE = objSelectedObject

End Function

Private Function GetAccessObjectOfVbComponente(ByVal vbc As VBComponent) As AccessObject

    If vbc.Type = vbext_ct_Document Then
        If Left(vbc.Name, 5) = "Form_" Then
            Set GetAccessObjectOfVbComponente = CurrentProject.AllForms(Mid(vbc.Name, Len("Form_") + 1))
        ElseIf Left(vbc.Name, 7) = "Report_" Then
            Set GetAccessObjectOfVbComponente = CurrentProject.AllReports(Mid(vbc.Name, Len("Report_") + 1))
        End If
    Else
        Set GetAccessObjectOfVbComponente = CurrentProject.AllModules(vbc.Name)
    End If

End Function

Private Function GetSelectedVBComponentsFromVbeProjectTreeGroup() As VBComponent()

    Dim SelectedProjectTreeviewItemText As String
    Dim SelectedProjectTreeviewItemLevel As Long
    Dim ExportAll As Boolean
    Dim arrVbc() As VBComponent
    Dim VbComponentTypeFilter As vbext_ComponentType
    Dim vbc As VBComponent
    Dim vbcCol As VBComponents
    Dim ItemCnt As Long

    With New clsVbeProjectWindowService
        SelectedProjectTreeviewItemText = .GetSelectedProjectExplorerItem(SelectedProjectTreeviewItemLevel)
    End With

    If SelectedProjectTreeviewItemLevel = 1 Then
        ExportAll = True
    Else
        Select Case SelectedProjectTreeviewItemText
            Case "Class Modules"
                VbComponentTypeFilter = vbext_ct_ClassModule
            Case "Modules"
                VbComponentTypeFilter = vbext_ct_StdModule
            Case "Microsoft Access Class Objects"
                VbComponentTypeFilter = vbext_ct_Document
            ' ... Forms

        End Select
    End If

    Set vbcCol = CurrentVBProject.VBComponents
    ReDim arrVbc(vbcCol.Count - 1)

    For Each vbc In vbcCol
        If ExportAll Or vbc.Type = VbComponentTypeFilter Then
            Set arrVbc(ItemCnt) = vbc
            ItemCnt = ItemCnt + 1
        End If
    Next

    If ItemCnt > 0 Then
        ReDim Preserve arrVbc(ItemCnt - 1)
        GetSelectedVBComponentsFromVbeProjectTreeGroup = arrVbc
    End If

End Function

Private Sub RunSelectedCommand(ByRef objSelected() As AccessObject, ByVal strCommandCode As String)

    If IsEmptyArray(objSelected) Then
        Exit Sub
    End If

    Select Case strCommandCode
        Case CompareSelectedWithBaseCommandCode
            CompareObjectWithBase objSelected

        Case ExportSelectedCommandCode
            ExportObjects objSelected

        Case LoadSelectedCommandCode
            LoadObjects objSelected

    End Select

End Sub

Private Sub CompareObjectWithBase(ByRef objToCompare() As AccessObject)

    Dim strSourceFilePath As String

    If objToCompare(0) Is Nothing Then
        Exit Sub
    End If

    modImportExport.ExportSingleObject objToCompare(0)

    With GetClassFromObject(objToCompare(0))
        strSourceFilePath = .SourceFile
        If .ComponentType = edbForm Then
            strSourceFilePath = Left(strSourceFilePath, Len(strSourceFilePath) - Len(".bas")) & ".cls"
        End If
    End With

    modObjects.Diff.Files strSourceFilePath, vbNullString

End Sub

Private Sub ExportObjects(ByRef objToExport() As Access.AccessObject)

    If LBound(objToExport) = UBound(objToExport) Then   ' SingleObject
        VCS.ExportObject objToExport(0)
        Exit Sub
    End If

    VCS.ExportObjects objToExport

End Sub

Private Sub LoadObjects(ByRef objToLoad() As AccessObject)
    VCS.LoadObjects objToLoad
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsBuild
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Build options section: build behavior, git integration, and build hooks.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection

' Overlay list for editing the procedures of the Run* options
Private m_ValueList As clsOptionValueListOverlay


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values into controls on initial display.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    LoadOptions
    InitValueList
End Sub


'---------------------------------------------------------------------------------------
' Procedure : InitValueList
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Prepare the overlay list used by the "..." buttons of the Run* options.
'---------------------------------------------------------------------------------------
'
Private Sub InitValueList()
    Set m_ValueList = New clsOptionValueListOverlay
    m_ValueList.Init Me, Me.sfrValueList
End Sub


Private Sub cmdRunAfterBuildList_Click()
    m_ValueList.OpenFor Me.txtRunAfterBuild, Me.cmdRunAfterBuildList
End Sub

Private Sub cmdRunBeforeMergeList_Click()
    m_ValueList.OpenFor Me.txtRunBeforeMerge, Me.cmdRunBeforeMergeList
End Sub

Private Sub cmdRunAfterMergeList_Click()
    m_ValueList.OpenFor Me.txtRunAfterMerge, Me.cmdRunAfterMergeList
End Sub

Private Sub sfrValueList_Exit(Cancel As Integer)
    m_ValueList.ApplyValues
End Sub

Private Sub Form_Timer()
    m_ValueList.HideList
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IOptionsSection_LoadOptions / IOptionsSection_SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : IOptionsSection interface implementation.
'---------------------------------------------------------------------------------------
'
Private Sub IOptionsSection_LoadOptions()
    LoadOptions
End Sub

Private Sub IOptionsSection_SaveOptions()
    SaveOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values from the working options instance into form controls.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()

    With Form_frmVCSOptions.OptionsInstance
        Me.chkExportAfterMerge = .ExportAfterMerge
        Me.chkForceImportOriginalQuerySQL = .ForceImportOriginalQuerySQL
        Me.chkMergeTableData = .MergeTableData
        Me.chkSkipReopenBeforeMerge = .SkipReopenBeforeMerge
        Me.txtRunBeforeBuild = .RunBeforeBuild
        Me.txtRunAfterBuild = .RunAfterBuild
        Me.txtValidateAfterBuild = .ValidateAfterBuild
        Me.txtRunBeforeMerge = .RunBeforeMerge
        Me.txtRunAfterMerge = .RunAfterMerge
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save form control values back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()

    With Form_frmVCSOptions.OptionsInstance
        .ExportAfterMerge = Nz(Me.chkExportAfterMerge)
        .ForceImportOriginalQuerySQL = Nz(Me.chkForceImportOriginalQuerySQL)
        .MergeTableData = Nz(Me.chkMergeTableData)
        .SkipReopenBeforeMerge = Nz(Me.chkSkipReopenBeforeMerge)
        .RunBeforeBuild = Nz(Me.txtRunBeforeBuild)
        .RunAfterBuild = Nz(Me.txtRunAfterBuild)
        .ValidateAfterBuild = Nz(Me.txtValidateAfterBuild)
        .RunBeforeMerge = Nz(Me.txtRunBeforeMerge)
        .RunAfterMerge = Nz(Me.txtRunAfterMerge)
    End With

End Sub

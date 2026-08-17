Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsExport
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Export options section: source file settings, sanitization, content
'           : options, printer settings detail, and export hooks.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection

Private WithEvents m_OptionValueList As Form_frmVCSOptionsValueList
Attribute m_OptionValueList.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values and populate combo box row sources.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim intSanitizeLevel As eSanitizeLevel
    Dim intEnvLevel As eUseEnvConnections
    Dim varFormat As Variant
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Populate sanitize level combo
    With Me.cboSanitizeLevel
        .RowSource = vbNullString
        For intSanitizeLevel = eslMinimal To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & opt.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Populate sanitize colors combo
    With Me.cboSanitizeColors
        .RowSource = vbNullString
        For intSanitizeLevel = eslNone To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & opt.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Populate .env connection string combo
    With Me.cboUseEnvForConnections
        .RowSource = vbNullString
        For intEnvLevel = uecAuto To eUseEnvConnections.[_Last] - 1
            .AddItem intEnvLevel & ";" & opt.GetUseEnvConnectionsName(intEnvLevel)
        Next intEnvLevel
        .ControlTipText = T("Auto: use .env when credentials are present; prompt for new on build. " & _
            "Always: store all connections in .env; save silently on build. " & _
            "Never: keep connection strings in source. " & _
            "Passwords never written to source (export format 5.0.0+).")
    End With

    ' Populate export format version combo
    With Me.cboExportFormatVersion
        .RowSource = vbNullString
        For Each varFormat In GetExportFormatVersions
            If varFormat = LatestExportFormat Then
                .AddItem varFormat & ";" & ExportFormatToVersion(CLng(varFormat)) & " (latest)"
            Else
                .AddItem varFormat & ";" & ExportFormatToVersion(CLng(varFormat))
            End If
        Next varFormat
    End With

    LoadOptions

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

    Dim varKey As Variant

    With Form_frmVCSOptions.OptionsInstance
        ' Source Files
        Me.chkUseFastSave = .UseFastSave
        Me.chkSplitLayoutFromVBA = .SplitLayoutFromVBA
        Me.cboExportFormatVersion = .ExportFormatVersion

        ' Sanitization
        Me.cboSanitizeLevel = .SanitizeLevel
        Me.cboSanitizeColors = .SanitizeColors

        ' Content
        Me.chkSaveQuerySQL = .SaveQuerySQL
        Me.chkFormatSQL = .FormatSQL
        Me.chkUseDeterministicQueryExport = .UseDeterministicQueryExport
        Me.chkSaveTableSQL = .SaveTableSQL
        Me.chkSavePrintVars = .SavePrintVars
        Me.chkDecodeConditionalFormatting = .DecodeConditionalFormatting
        Me.chkExtractThemeFiles = .ExtractThemeFiles
        Me.cboUseEnvForConnections = .UseEnvForConnections

        ' Printer settings detail
        For Each varKey In .ExportPrintSettings.Keys
            Me.Controls("chk" & varKey) = .ExportPrintSettings(varKey)
        Next varKey

        ' Hooks
        Me.txtRunBeforeExport = .RunBeforeExport
        Me.txtRunAfterExport = .RunAfterExport
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

    Dim varKey As Variant
    Dim dSettings As Dictionary

    With Form_frmVCSOptions.OptionsInstance
        ' Source Files
        .UseFastSave = Nz(Me.chkUseFastSave)
        .ExportFormatVersion = Nz(Me.cboExportFormatVersion)

        ' Check for SplitLayoutFromVBA change before saving
        OnSplitLayoutChange Nz(Me.chkSplitLayoutFromVBA)
        .SplitLayoutFromVBA = Nz(Me.chkSplitLayoutFromVBA)

        ' Sanitization
        .SanitizeLevel = Nz(Me.cboSanitizeLevel)
        .SanitizeColors = Nz(Me.cboSanitizeColors)

        ' Content
        .SaveQuerySQL = Nz(Me.chkSaveQuerySQL)
        .FormatSQL = Nz(Me.chkFormatSQL)
        .UseDeterministicQueryExport = Nz(Me.chkUseDeterministicQueryExport)
        .SaveTableSQL = Nz(Me.chkSaveTableSQL)
        .SavePrintVars = Nz(Me.chkSavePrintVars)
        .DecodeConditionalFormatting = Nz(Me.chkDecodeConditionalFormatting)
        .ExtractThemeFiles = Nz(Me.chkExtractThemeFiles)
        .UseEnvForConnections = Nz(Me.cboUseEnvForConnections)

        ' Printer settings detail
        Set dSettings = .ExportPrintSettings
        For Each varKey In dSettings.Keys
            dSettings(varKey) = CBool(Me.Controls("chk" & varKey))
        Next varKey

        ' Hooks
        .RunBeforeExport = Nz(Me.txtRunBeforeExport)
        .RunAfterExport = Nz(Me.txtRunAfterExport)
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OnSplitLayoutChange
' Author    : Adam Waller
' Date      : 11/9/2023
' Purpose   : When the user enables Split Layout from VBA, suggest the wiki page
'           : about preserving Git history for split files.
'---------------------------------------------------------------------------------------
'
Private Sub OnSplitLayoutChange(varNewValue As Variant)

    With Form_frmVCSOptions.OptionsInstance
        If .SplitLayoutFromVBA = varNewValue Then Exit Sub
    End With

    If varNewValue = True Then
        If Git.Installed Then
            If Git.IsInsideRepository Then
                If MsgBox2(T("May I make a Suggestion?"), _
                    T("This project appears to be within a Git repository. This add-in includes a special utility " & _
                    "that can split the files (layout and VBA) while preserving this history of previous changes in BOTH files."), _
                    T("Would you like to see additional information on this from the wiki?"), vbQuestion + vbYesNo) = vbYes Then
                    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Split-Files"
                End If
            End If
        End If
    End If

End Sub

'---------------------------------------------------------------------------------------
' Option value list procedures:
'
Private Sub cmdRunBeforeExportOpenValueList_Click()
    UseOptionValueListForTextbox Me.txtRunBeforeExport, Me.cmdRunBeforeExportOpenValueList, Me.sfrExportOptionsValueList
End Sub

Private Sub cmdRunAfterExportOpenValueList_Click()
    UseOptionValueListForTextbox Me.txtRunAfterExport, Me.cmdRunAfterExportOpenValueList, Me.sfrExportOptionsValueList
End Sub

Private Sub UseOptionValueListForTextbox(ByVal OptionTextBox As TextBox, ByVal OpenListCommandButton As CommandButton, OptionsValueListSubForm As SubForm, Optional ByVal ListHeight As Long = 2000)

    Dim OptionValues() As String
    Dim i As Long

    If Not (m_OptionValueList Is Nothing) Then
        DoEvents
    End If

    OptionValues = Split(Nz(OptionTextBox.Value, vbNullString), ":")
    If Len(OptionTextBox.Value) > 0 Then
        For i = LBound(OptionValues) To UBound(OptionValues)
            OptionValues(i) = Trim(OptionValues(i))
        Next i
    End If

    With OptionsValueListSubForm
        .Left = OptionTextBox.Left
        .Width = OptionTextBox.Width + OpenListCommandButton.Width
        .Top = OptionTextBox.Top + OptionTextBox.Height - ListHeight
        .Height = ListHeight
        .Form.OptionValueListCode = OptionTextBox.Name
        .Form.SetData OptionValues
        .Enabled = True
        .Visible = True
        .SetFocus
    End With

    Set m_OptionValueList = Me.sfrExportOptionsValueList.Form

End Sub

Private Sub m_OptionValueList_CloseOptionList()
    With Me.sfrExportOptionsValueList.Form
        Me.Controls(.OptionValueListCode).SetFocus ' => sfrExportOptionsValueList_Exit
    End With
End Sub

Private Sub sfrExportOptionsValueList_Exit(Cancel As Integer)
    SaveOptionValues
    Me.TimerInterval = 1
End Sub

Private Sub SaveOptionValues()
    With Me.sfrExportOptionsValueList.Form
        Me.Controls(.OptionValueListCode).Value = Join(.GetData, " : ")
    End With
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    With Me.sfrExportOptionsValueList
        .Enabled = False
        .Visible = False
        If Me.ActiveControl.Name = .Form.OptionValueListCode Then
            With Me.ActiveControl
                .SelStart = Len(.Text)
            End With
        End If
    End With
End Sub

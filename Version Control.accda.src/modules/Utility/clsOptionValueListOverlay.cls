VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptionValueListOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOptionValueListOverlay
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Shows a frmVCSOptionsValueList subform as an overlay above an option text
'           : box so the entries of a Run* option (separated by colons) can be edited
'           : one per row. Used by the options forms; the host form forwards its
'           : events to OpenFor, ApplyValues (subform Exit) and HideList (Timer).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

' Height of the overlay list (twips)
Private Const LIST_HEIGHT As Long = 2000

' Separator used when the rows are joined back into the option value
Private Const VALUE_SEPARATOR As String = " : "

Private m_frmHost As Form
Private m_sfrList As SubForm
Private WithEvents m_frmList As Form_frmVCSOptionsValueList
Attribute m_frmList.VB_VarHelpID = -1
Private m_txtTarget As TextBox
Private m_blnCancel As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Connect the overlay to the host form and its (hidden) subform control.
'---------------------------------------------------------------------------------------
'
Public Sub Init(ByVal frmHost As Form, ByVal sfrList As SubForm)
    Set m_frmHost = frmHost
    Set m_sfrList = sfrList
    Set m_frmList = sfrList.Form
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenFor
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Show the list for txtTarget above the text box, as wide as the text box
'           : and its button.
'---------------------------------------------------------------------------------------
'
Public Sub OpenFor(ByVal txtTarget As TextBox, ByVal cmdOpen As CommandButton)

    Dim lngTop As Long

    ' Discard a pending hide of the previous list, it would hide this one
    m_frmHost.TimerInterval = 0

    Set m_txtTarget = txtTarget
    m_blnCancel = False
    m_frmList.SetValues SplitValues(Nz(txtTarget.Value, vbNullString))

    ' Escape must only close the list, not the options dialog
    Form_frmVCSOptions.SetCancelOnEscape False

    lngTop = txtTarget.Top + txtTarget.Height - LIST_HEIGHT
    If lngTop < 0 Then lngTop = 0

    With m_sfrList
        .Left = txtTarget.Left
        .Width = cmdOpen.Left + cmdOpen.Width - txtTarget.Left
        .Top = lngTop
        .Height = LIST_HEIGHT
        .Enabled = True
        .Visible = True
        .SetFocus
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyValues
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Write the rows back into the text box (call from the subform's Exit
'           : event). The list is hidden by HideList, because a control cannot be
'           : hidden while it still has the focus.
'---------------------------------------------------------------------------------------
'
Public Sub ApplyValues()
    If m_txtTarget Is Nothing Then Exit Sub
    If m_blnCancel Then
        m_blnCancel = False
    Else
        m_txtTarget.Value = JoinValues(m_frmList.GetValues)
    End If
    m_frmHost.TimerInterval = 1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HideList
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Hide the list (call from the host form's Timer event).
'---------------------------------------------------------------------------------------
'
Public Sub HideList()

    m_frmHost.TimerInterval = 0
    With m_sfrList
        .Enabled = False
        .Visible = False
    End With

    If Not m_txtTarget Is Nothing Then
        If IsActiveControl(m_txtTarget) Then m_txtTarget.SelStart = Len(m_txtTarget.Text)
    End If
    Set m_txtTarget = Nothing

    Form_frmVCSOptions.SetCancelOnEscape True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseList
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Close the list if it is open, so its values are written back into the
'           : text box (call before reading the option controls, e.g. when saving).
'---------------------------------------------------------------------------------------
'
Public Sub CloseList()
    If m_sfrList.Visible Then MoveFocusToTarget
End Sub


Private Sub m_frmList_CloseRequested()
    MoveFocusToTarget
End Sub


Private Sub m_frmList_CancelRequested()
    CancelList
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CancelList
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Close the list without writing the values back into the text box.
'---------------------------------------------------------------------------------------
'
Private Sub CancelList()
    m_blnCancel = True
    MoveFocusToTarget
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MoveFocusToTarget
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Leave the list (triggers the subform's Exit event and so ApplyValues).
'---------------------------------------------------------------------------------------
'
Private Sub MoveFocusToTarget()
    If Not m_txtTarget Is Nothing Then m_txtTarget.SetFocus
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsActiveControl
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : True if ctl has the focus on the host form. (ActiveControl raises an
'           : error when no control has the focus.)
'---------------------------------------------------------------------------------------
'
Private Function IsActiveControl(ByVal ctl As Control) As Boolean
    LogUnhandledErrors
    On Error Resume Next
    IsActiveControl = (m_frmHost.ActiveControl.Name = ctl.Name)
    If Err Then Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : SplitValues / JoinValues
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Convert between the option value and the list rows with the same rules
'           : as RunExternalProcedure (clsProcedureRunner.SplitProcedureList).
'---------------------------------------------------------------------------------------
'
Private Function SplitValues(ByVal strOptionValue As String) As Collection
    With New clsProcedureRunner
        Set SplitValues = .SplitProcedureList(strOptionValue)
    End With
End Function

Private Function JoinValues(ByVal colValues As Collection) As String

    Dim varValue As Variant

    With New clsConcat
        .AppendOnAdd = VALUE_SEPARATOR
        For Each varValue In colValues
            .Add CStr(varValue)
        Next varValue
        If colValues.Count > 0 Then .Remove Len(VALUE_SEPARATOR)
        JoinValues = .GetStr
    End With

End Function

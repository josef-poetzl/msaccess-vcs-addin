Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsValueList
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Continuous list used as an overlay on the options forms to edit the
'           : entries of a multi-value option (several procedures of a Run* option).
'           : The values live in a disconnected ADODB recordset; empty rows are
'           : dropped by GetValues. Positioning and saving are done by
'           : clsOptionValueListOverlay.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Event CloseRequested()

Private Const VALUE_FIELD As String = "OptionValue"
Private Const VALUE_FIELD_SIZE As Long = 1024

Private m_rsValues As ADODB.Recordset


'---------------------------------------------------------------------------------------
' Procedure : SetValues
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Show the given values (strings) as list rows.
'---------------------------------------------------------------------------------------
'
Public Sub SetValues(ByVal colValues As Collection)
    Set m_rsValues = NewValueRecordset(colValues)
    Set Me.Recordset = m_rsValues
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetValues
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Return the non-empty values of all rows in list order.
'---------------------------------------------------------------------------------------
'
Public Function GetValues() As Collection

    Dim colValues As Collection
    Dim rsClone As ADODB.Recordset
    Dim strValue As String

    Set colValues = New Collection
    If Not m_rsValues Is Nothing Then
        Set rsClone = m_rsValues.Clone
        If rsClone.RecordCount > 0 Then
            rsClone.MoveFirst
            Do While Not rsClone.EOF
                strValue = Trim$(Nz(rsClone.Fields(VALUE_FIELD).Value, vbNullString))
                If Len(strValue) Then colValues.Add strValue
                rsClone.MoveNext
            Loop
        End If
        rsClone.Close
        Set rsClone = Nothing
    End If

    Set GetValues = colValues

End Function


'---------------------------------------------------------------------------------------
' Procedure : NewValueRecordset
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Create a disconnected, updatable recordset with one value field.
'---------------------------------------------------------------------------------------
'
Private Function NewValueRecordset(ByVal colValues As Collection) As ADODB.Recordset

    Dim rsValues As ADODB.Recordset
    Dim varValue As Variant

    Set rsValues = New ADODB.Recordset
    With rsValues
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append VALUE_FIELD, adVarWChar, VALUE_FIELD_SIZE, adFldUpdatable + adFldIsNullable
        .Open
        For Each varValue In colValues
            .AddNew
            .Fields(VALUE_FIELD).Value = varValue
            .Update
        Next varValue
        If .RecordCount > 0 Then .MoveFirst
    End With

    Set NewValueRecordset = rsValues

End Function


Private Sub Form_Current()
    PrepareNewRowInput
End Sub


Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    ' Delete the row without the Access confirmation prompt
    Response = acDataErrContinue
End Sub


Private Sub txtValueInsert_Change()
    StartNewRowFromInput
End Sub


Private Sub cmdClose_Click()
    RaiseEvent CloseRequested
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PrepareNewRowInput
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : On the new-record row, route typing into the unbound txtValueInsert
'           : (the ADODB-bound row cannot start a new record by itself).
'---------------------------------------------------------------------------------------
'
Private Sub PrepareNewRowInput()
    If Me.NewRecord Then
        With Me.txtValueInsert
            .Value = Me.txtValue.Value
            .SetFocus
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StartNewRowFromInput
' Author    : Josef Pötzl
' Date      : 10/20/2026
' Purpose   : Add a recordset row for the first typed character and continue editing
'           : in the bound text box.
'---------------------------------------------------------------------------------------
'
Private Sub StartNewRowFromInput()
    If Me.NewRecord Then
        m_rsValues.AddNew
        m_rsValues.Fields(VALUE_FIELD).Value = Me.txtValueInsert.Text
        Me.Recordset.Bookmark = m_rsValues.Bookmark

        Me.txtValue.SetFocus
        Me.txtValue.Text = Me.txtValueInsert.Value
        Me.txtValue.SelStart = Len(Me.txtValue.Text)
    End If
End Sub

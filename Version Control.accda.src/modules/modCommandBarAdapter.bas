Attribute VB_Name = "modCommandBarAdapter"
Option Compare Database
Option Explicit

Private m_CommandBarAdapter As clsCommandBarAdapter

Public Sub InitCommandBarAdapter(Optional ByVal ForceNewInstance As Boolean = False)

    If ForceNewInstance Then
        DisposeCommandBarAdapter
    End If
    If m_CommandBarAdapter Is Nothing Then
        Set m_CommandBarAdapter = New clsCommandBarAdapter
    End If

End Sub

Public Sub DisposeCommandBarAdapter()
    Set m_CommandBarAdapter = Nothing
End Sub

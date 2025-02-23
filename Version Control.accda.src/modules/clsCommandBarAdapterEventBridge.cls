VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandBarAdapterEventBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event Clicked(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter)

Public Sub RaiseClicked(ByVal objCommandBarButtonAdapter As clsCommandBarButtonAdapter)
    RaiseEvent Clicked(objCommandBarButtonAdapter)
End Sub

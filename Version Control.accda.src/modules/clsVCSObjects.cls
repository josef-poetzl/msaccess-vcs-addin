VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : modObjects
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Wrapper functions for classes and other objects available globally.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


Public Perf As clsPerformance
Public Log As clsLog
Public Options As clsOptions
Public VCSIndex As clsVCSIndex
Public Worker As clsWorker
Public Git As clsGitIntegration
Public Translation As clsTranslation

    ' Keep a persistent reference to file system object after initializing version control.
    ' This way we don't have to recreate this object dozens of times while using VCS.
Public FSO As Scripting.FileSystemObject

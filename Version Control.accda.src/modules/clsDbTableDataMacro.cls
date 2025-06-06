﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTableDataMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_Table As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strHash As String

    ' Make sure the table exists, and is accessible
    If Not TableExists(m_Table.Name) Then
        Log.Error eelError, "Unable to access table " & m_Table.Name & _
            ". Please ensure that this table exists and is accessible in the current database.", _
            ModuleName(Me) & ".Export"
        Exit Sub
    End If

    ' We THINK that this table probably contains a table macro, but just in case it
    ' doesn't, we have some error handling built into the export function below.

    ' Save structure in XML format
    strHash = SaveComponentAsText(acTableDataMacro, m_Table.Name, Nz2(strAlternatePath, IDbComponent_SourceFile))

    ' Save index if we found an item to export
    If strHash <> vbNullString Then
        VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), strHash
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim blnImportCheck As Boolean
    Dim strName As String

    ' Only import files with the correct extension.
    If Not strFile Like "*.xml" Then Exit Sub

    strName = GetObjectNameFromFileName(strFile)
    blnImportCheck = LoadComponentFromText(acTableDataMacro, strName, strFile)
    If Not blnImportCheck Or Log.ErrorLevel = eelCritical Then Exit Sub
    ' Update index
    Set m_Table = CurrentData.AllTables(strName)
    VCSIndex.Update Me, eatImport, GetFileHash(strFile)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)

    Dim strTemp As String

    If FSO.FileExists(strFile) Then
        ' Load the table data macro into the existing table, overwriting any existing macro
        IDbComponent_Import strFile
    Else
        ' Create and load a blank file to clear any existing table data macros.
        ' (Tested in Access 2010 with no errors on 6/15/2022)
        strTemp = GetTempFile   ' Create an empty temporary file.
        LoadFromText acTableDataMacro, GetObjectNameFromFileName(strFile), strTemp
        DeleteFile strTemp

        ' Remove entry from index
        VCSIndex.Remove Me, strFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'           : Credit to StackOverflow for the clue on how to determine if a table has
'           : a data macro. https://stackoverflow.com/questions/31755802/
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim dbs As Database
    Dim tdf As TableDef
    Dim cTable As IDbComponent
    Dim strSql As String
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        Set dbs = CurrentDb
        For Each tdf In dbs.TableDefs
            ' Skip system, temp, and linked tables
            If Left$(tdf.Name, 4) <> "MSys" Then
                If Left$(tdf.Name, 1) <> "~" Then
                    If Len(tdf.Connect) = 0 Then
                        ' Check to see if the table has a data macro
                        strSql = "(LvExtra is not null) and (Type = 1) and ([Name] = """ & tdf.Name & """)"
                        If DCount("[Name]", "MSysObjects", strSql) > 0 Then
                            Set cTable = New clsDbTableDataMacro
                            Set cTable.DbObject = CurrentData.AllTables(tdf.Name)
                            blnAdd = True
                            If blnModifiedOnly Then blnAdd = cTable.IsModified
                            If blnAdd Then m_Items(blnModifiedOnly).Add cTable.SourceFile, cTable
                        End If
                    End If
                End If
            End If
        Next tdf

        Set tdf = Nothing
        Set dbs = Nothing

    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.xml")
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    ' A table modification doesn't always mean the macro was modified, but we will
    ' export on any table modification, just to be on the safe side.
    IDbComponent_IsModified = (m_Table.DateModified > VCSIndex.Item(Me).ExportDate)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    If m_Table Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Table.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Table Data Macros"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "tdmacros" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "xml"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_Name = m_Table.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & ".xml"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    IDbComponent_QuickCount = 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbTableDataMacro
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Table
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Table = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property

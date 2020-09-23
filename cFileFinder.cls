VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17440 + vbObjectError

Public Enum EErrFileFinder
    eErrFileFinder_NotInitialized = E_ERR_BASE + 1
    eErrFileFinder_InvalidFilePattern
    eErrFileFinder_InvalidStartPath
    eErrFileFinder_InvalidOutputSytle
    eErrFileFinder_SubscriptOutOfRange
    eErrFileFinder_FilesCollectionNotAvailable
    eErrFileFinder_FindFilesNotCalledOrInProgress
    eErrFileFinder_ComponentFailure
End Enum

Private Const S_ERR_NotInitialized = "Object not properly initialized. Set the FilePattern and StartPath properties first"
Private Const S_ERR_InvalidFilePattern = "Invalid file pattern"
Private Const S_ERR_InvalidStartPath = "Invalid start path"
Private Const S_ERR_InvalidOutputStyle = "Invalid output style"
Private Const S_ERR_SubscriptOutOfRange = "Subscript out of range"
Private Const S_ERR_FilesCollectionNotAvailable = "Files collection not available for the current output style"
Private Const S_ERR_FindFilesNotCalledOrInProgress = "FindFiles was not called or is still executing"
Private Const S_ERR_ComponentFailure = "CFileFinder component failure"

' Public Class Enums
Public Enum EFileFinderOutputStyle
    eFileFinderOutputStyle_UseEvents
    eFileFinderOutputStyle_FillFilesCollection
End Enum

' Private class constants
Private Const MAX_LENGTH = 512

' Private class type definitions
Private Type FILETIME
    dwLowDateTime        As Long
    dwHighDateTime       As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes     As Long
    ftCreationTime       As FILETIME
    ftLastAccessTime     As FILETIME
    ftLastWriteTime      As FILETIME
    nFileSizeHigh        As Long
    nFileSizeLow         As Long
    dwReserved0          As Long
    dwReserved1          As Long
    cFileName            As String * MAX_LENGTH
    cAlternate           As String * 14
End Type

' Private class API function declarations
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' Private variables for internal class use
Private m_bDone          As Boolean
Private m_colFiles       As Collection
Private m_iRecCount      As Long

' Private variables to hold property values
Private m_bIncludeSubfolders     As Boolean
Private m_iCount                 As Long
Private m_osOutputStyle          As EFileFinderOutputStyle
Private m_sStartPath             As String
Private m_sFilePattern           As String

' Event definitions
Public Event FoundFile(FileName As String)
Public Event SearchingDirectory(Path As String)
Public Event Finished()


'*****************************************************************************************
'* Sub         : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Sub Init(FilePattern As String, StartPath As String, Optional IncludeSubfolders As Boolean = False, Optional OutputStyle As EFileFinderOutputStyle = eFileFinderOutputStyle_UseEvents)
    On Error GoTo hComponentFailure

    If Len(FilePattern) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidFilePattern, App.EXEName & ".CFileFinder", S_ERR_InvalidFilePattern
    End If

    m_sFilePattern = FilePattern

    If Len(StartPath) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidStartPath, App.EXEName & ".CFileFinder", S_ERR_InvalidStartPath
    End If

    m_sStartPath = StartPath

    m_bIncludeSubfolders = IncludeSubfolders

    If Not ((OutputStyle = eFileFinderOutputStyle_FillFilesCollection) Or (OutputStyle = eFileFinderOutputStyle_UseEvents)) Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidOutputSytle, App.EXEName & ".CFileFinder", S_ERR_InvalidOutputStyle
    End If

    m_osOutputStyle = OutputStyle

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Sub


'*****************************************************************************************
'* Property    : Count
'* Notes       : Returns a long integer value containing the number of files that matched
'*               the search criteria.
'*****************************************************************************************
Public Property Get Count() As Long
    On Error GoTo hComponentFailure

    If Not m_bDone Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_FindFilesNotCalledOrInProgress, App.EXEName & ".CFileFinder", S_ERR_FindFilesNotCalledOrInProgress
    End If

    Count = m_iCount

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Property    : FilePattern
'* Notes       : Returns or sets a string value containing the name of the file to be
'*               searched. It can contains wildcard characters (e.g. "*.txt").
'*****************************************************************************************
Public Property Let FilePattern(Value As String)
    On Error GoTo hComponentFailure

    If Len(FilePattern) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidFilePattern, App.EXEName & ".CFileFinder", S_ERR_InvalidFilePattern
    End If

    If m_sFilePattern <> Value Then
        m_bDone = False
        m_sFilePattern = Value
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property

Public Property Get FilePattern() As String
    On Error GoTo hComponentFailure

    FilePattern = m_sFilePattern

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Property    : Files
'* Notes       : Returns a string value containing the name of the file with the specified
'*               index in the founded files array.
'*****************************************************************************************
Public Property Get Files(Index As Long) As String
    On Error GoTo hComponentFailure

    If Not m_bDone Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_FindFilesNotCalledOrInProgress, App.EXEName & ".CFileFinder", S_ERR_FindFilesNotCalledOrInProgress
    End If

    If m_osOutputStyle = eFileFinderOutputStyle_UseEvents Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_FilesCollectionNotAvailable, App.EXEName & ".CFileFinder", S_ERR_FilesCollectionNotAvailable
    End If

    If (Index < 1) Or (Index > m_iCount) Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_SubscriptOutOfRange, App.EXEName & ".CFileFinder", S_ERR_SubscriptOutOfRange
    End If

    Files = CStr(m_colFiles(Index))

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Property    : IncludeSubfolders
'* Notes       : Returns or sets if the search operation will be perform in all the
'*               subfolders of the starting path.
'*****************************************************************************************
Public Property Let IncludeSubfolders(Value As Boolean)
    On Error GoTo hComponentFailure

    If m_bIncludeSubfolders <> Value Then
        m_bDone = False
        m_bIncludeSubfolders = Value
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property

Public Property Get IncludeSubfolders() As Boolean
    On Error GoTo hComponentFailure

    IncludeSubfolders = m_bIncludeSubfolders

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Property    : OutputStyle
'* Notes       : Returns or sets the search operation output style.
'*****************************************************************************************
Public Property Let OutputStyle(Value As EFileFinderOutputStyle)
    On Error GoTo hComponentFailure

    If Not ((OutputStyle = eFileFinderOutputStyle_FillFilesCollection) Or (OutputStyle = eFileFinderOutputStyle_UseEvents)) Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidOutputSytle, App.EXEName & ".CFileFinder", S_ERR_InvalidOutputStyle
    End If

    If m_osOutputStyle <> Value Then
        m_bDone = False
        m_osOutputStyle = Value
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property

Public Property Get OutputStyle() As EFileFinderOutputStyle
    On Error GoTo hComponentFailure

    OutputStyle = m_osOutputStyle

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Property    : StartPath
'* Notes       : Returns or sets a string value containing the starting path for the
'*               search operation.
'*****************************************************************************************
Public Property Let StartPath(Value As String)
    On Error GoTo hComponentFailure

    If Len(StartPath) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_InvalidStartPath, App.EXEName & ".CFileFinder", S_ERR_InvalidStartPath
    End If

    If m_sStartPath <> Value Then
        m_bDone = False
        m_sStartPath = Value
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property

Public Property Get StartPath() As String
    On Error GoTo hComponentFailure

    StartPath = m_sStartPath

    Exit Property

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Property


'*****************************************************************************************
'* Sub         : FindFiles
'* Notes       : Searches for the files that match the specified file pattern starting
'*               with the path specified by the StartPath property.
'*****************************************************************************************
Public Sub FindFiles()
    On Error GoTo hComponentFailure

    If (Len(m_sFilePattern) = 0) Or (Len(m_sStartPath) = 0) Then
        On Error GoTo 0
        Err.Raise eErrFileFinder_NotInitialized, App.EXEName & ".CFileFinder", S_ERR_NotInitialized
    End If

    m_iCount = 0
    Set m_colFiles = Nothing

    If m_osOutputStyle = eFileFinderOutputStyle_UseEvents Then
        m_iRecCount = 0
        Call FindFilesInFolderAsync(m_sStartPath)
    Else
        Set m_colFiles = New Collection
        Call FindFilesInFolderSync(m_sStartPath)
        m_bDone = True
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Sub


'*****************************************************************************************
'* Sub         : FindFilesInFolderAsync
'* Notes       : Private routine to search for the files that match the specified file
'*               pattern. This routine generates the following events:
'*               - FoundFile (when a file that match the pattern id found)
'*               - SearchingDirectory (when the searched path has changed)
'*               - Finished (when the finished operation ends)
'*****************************************************************************************
Private Sub FindFilesInFolderAsync(Folder As String)
    On Error GoTo hComponentFailure

    Dim hFile   As Long
    Dim wfd     As WIN32_FIND_DATA
    Dim sName   As String

    m_iRecCount = m_iRecCount + 1

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"

    RaiseEvent SearchingDirectory(Folder)

    hFile& = FindFirstFile(Folder & m_sFilePattern, wfd)
    If hFile& > 0 Then
        m_iCount = m_iCount + 1
        RaiseEvent FoundFile(Folder & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1))

        Do While FindNextFile(hFile, wfd)
            m_iCount = m_iCount + 1
            RaiseEvent FoundFile(Folder & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1))
        Loop
    End If

    FindClose hFile

    If m_bIncludeSubfolders Then

        hFile& = FindFirstFile(Folder & "*.*", wfd)
        If hFile& > 0 Then

            Do While FindNextFile(hFile, wfd)

                sName = Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
                If Left$(sName, 1) <> "." Then

                    If wfd.dwFileAttributes And vbDirectory Then
                        FindFilesInFolderAsync Folder & sName
                        m_iRecCount = m_iRecCount - 1
                    End If

                End If

            Loop

        End If

        FindClose hFile

    End If

    If m_iRecCount = 1 Then
        m_bDone = True
        RaiseEvent Finished
        m_iRecCount = 0
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Sub


'*****************************************************************************************
'* Sub         : FindFilesInFolderSync
'* Notes       : Private routine to search for the files that match the specified file
'*               pattern. It generates an array which will hold the matching files.
'*****************************************************************************************
Private Sub FindFilesInFolderSync(Folder As String)
    On Error GoTo hComponentFailure

    Dim hFile           As Long
    Dim sName           As String
    Dim wfd             As WIN32_FIND_DATA

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"

    hFile& = FindFirstFile(Folder & m_sFilePattern, wfd)

    If hFile& > 0 Then
        m_iCount = m_iCount + 1
        m_colFiles.Add Folder & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)

        Do While FindNextFile(hFile, wfd)
            m_iCount = m_iCount + 1
            m_colFiles.Add Folder & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
        Loop

    End If

    FindClose hFile

    If m_bIncludeSubfolders Then

        hFile& = FindFirstFile(Folder & "*.*", wfd)

        If hFile& > 0 Then

            Do While FindNextFile(hFile, wfd)

                sName = Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
                If Left$(sName, 1) <> "." Then
                    If wfd.dwFileAttributes And vbDirectory Then
                        FindFilesInFolderSync Folder & sName
                    End If
                End If

            Loop

        End If

        FindClose hFile

    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Sub


'*****************************************************************************************
'* Sub         : Class_Initialize
'* Notes       : Class data space initialization.
'*****************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    m_osOutputStyle = eFileFinderOutputStyle_UseEvents

    Exit Sub
hComponentFailure:
    Err.Raise eErrFileFinder_ComponentFailure, App.EXEName & ".CFileFinder", S_ERR_ComponentFailure

End Sub


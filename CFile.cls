VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17420 + vbObjectError

Public Enum EErrFile
    eErrFile_NotInitialized = E_ERR_BASE + 1
    eErrFile_InvalidFileName
    eErrFile_FileNotFound
    eErrFile_CopyFailed
    eErrFile_MoveFailed
    eErrFile_DeleteFailed
    eErrFile_ComponentFailure
End Enum

Private Const S_ERR_NotInitialized = "Object not properly initialized. Set the PathFileName property to a non-empty value first"
Private Const S_ERR_InvalidFileName = "Invalid file name"
Private Const S_ERR_FileNotFound = "File not found"
Private Const S_ERR_CopyFailed = "File copy failed"
Private Const S_ERR_DeleteFailed = "File delete failed"
Private Const S_ERR_MoveFailed = "File move failed"
Private Const S_ERR_ComponentFailure = "CFile component failure"

' Public class enums
Public Enum EFileAttributes
    eFileAttribute_ReadOnly = &H1
    eFileAttribute_Hidden = &H2
    eFileAttribute_System = &H4
    eFileAttribute_Directory = &H10
    eFileAttribute_Archive = &H20
    eFileAttribute_Encrypted = &H40
    eFileAttribute_Normal = &H80
    eFileAttribute_Temporary = &H100
    eFileAttribute_SparseFile = &H200
    eFileAttribute_ReparsePoint = &H400
    eFileAttribute_Compressed = &H800
    eFileAttribute_Offline = &H1000
End Enum

' Private class constants
Private Const INVALID_HANDLE_VALUE = -1
Private Const CHUNKSIZE As Long = 16384
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
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpExistingFileName As String) As Long
Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' Private variables to hold property values
Private m_sPathFileName  As String


'*****************************************************************************************
'* Sub         : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Sub Init(PathFileName As String)
    On Error GoTo hComponentFailure
    
    If Len(PathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_InvalidFileName, App.EXEName & ".CFile", S_ERR_InvalidFileName
    End If

    m_sPathFileName = PathFileName

    Exit Sub

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Property    : Attributes
'* Notes       : Returns the attributes for the current specified file.
'*****************************************************************************************
Public Property Get Attributes() As EFileAttributes
    On Error GoTo hComponentFailure
    
    Dim wfd As WIN32_FIND_DATA
    Dim hFile As Long
    
    Attributes = 0
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    hFile = FindFirstFile(m_sPathFileName, wfd)
    
    If (hFile <> INVALID_HANDLE_VALUE) Then
        Attributes = wfd.dwFileAttributes
        
        Call FindClose(hFile)
    Else
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : Contents
'* Notes       : Returns a string value filled with the contents of the current specified
'*               file.
'*               Sets the contents of the file to the specified string.
'*****************************************************************************************
Public Property Let Contents(Value As String)
    On Error GoTo hComponentFailure
    
    Dim iFF As Long
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    iFF = FreeFile
    Open m_sPathFileName For Binary As #iFF
    
    Print #iFF, , Value
    
    Close #iFF

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property

Public Property Get Contents() As String

    On Error GoTo hComponentFailure
    
    Dim iFF As Long
    
    Contents = ""
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    iFF = FreeFile
    Open m_sPathFileName For Input As #iFF
    
    Contents = StrConv(InputB(LOF(iFF), iFF), vbUnicode)
    
    Close #iFF

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : FileName
'* Notes       : Returns a string value containing the file name (without path and
'*               extension) for the currently specified file.
'*****************************************************************************************
Public Property Get FileName() As String
    On Error GoTo hComponentFailure
    
    Dim iPos%, iSearchedPos%
    Dim sFileWithExt$
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    iPos = InStr(1, m_sPathFileName, "\")
    Do While iPos <> 0
        iSearchedPos = iPos
        iPos = InStr(iPos + 1, m_sPathFileName, "\")
    Loop
    
    sFileWithExt = Right$(m_sPathFileName, Len(m_sPathFileName) - iSearchedPos)
    
    iSearchedPos = 0
    
    iPos = InStr(1, sFileWithExt, ".")
    Do While iPos <> 0
        iSearchedPos = iPos
        iPos = InStr(iPos + 1, sFileWithExt, ".")
    Loop
    
    If iSearchedPos > 0 Then
        FileName = Left$(sFileWithExt, iSearchedPos - 1)
    Else
        FileName = sFileWithExt
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : FileExt
'* Notes       : Returns a string value containing the extension corresponding to the
'*               currently specified file.
'*****************************************************************************************
Public Property Get FileExt() As String
    On Error GoTo hComponentFailure
    
    Dim iPos As Integer
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    iPos = InStrRev(m_sPathFileName, ".")
    
    If iPos > 0 Then
        FileExt = Mid$(m_sPathFileName, iPos + 1)
    Else
        FileExt = ""
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : FilePath
'* Notes       : Returns a string value containing the full path for the current specified
'*               file. (e.g. for the "C:\Windows\System\user32.dll" file the returned
'*               value is "C:\Windows\System").
'*****************************************************************************************
Public Property Get FilePath() As String
    On Error GoTo hComponentFailure
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    FilePath = Left$(m_sPathFileName, Len(m_sPathFileName) - Len(FileName) - Len(FileExt) - 1)

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : PathFileName
'* Notes       : Returns or sets the current file (path included).
'*****************************************************************************************
Public Property Let PathFileName(Value As String)
    On Error GoTo hComponentFailure
        
    If Len(Value) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_InvalidFileName, App.EXEName & ".CFile", S_ERR_InvalidFileName
    End If
    
    m_sPathFileName = Value
    
    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property

Public Property Get PathFileName() As String
    On Error GoTo hComponentFailure
    
    PathFileName = m_sPathFileName
    
    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : ReadOnly
'* Notes       : Returns True if the current file is read-only, False otherwise.
'*****************************************************************************************
Public Property Get ReadOnly() As Boolean
    On Error GoTo hComponentFailure
    
    Dim wfd As WIN32_FIND_DATA
    Dim hFile As Long
    
    ReadOnly = False
   
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    hFile = FindFirstFile(m_sPathFileName, wfd)
    If (hFile <> INVALID_HANDLE_VALUE) Then
        If (wfd.dwFileAttributes And eFileAttribute_ReadOnly) = eFileAttribute_ReadOnly Then
            ReadOnly = True
        Else
            ReadOnly = False
        End If
        Call FindClose(hFile)
        
    Else
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : ShortFileName
'* Notes       : Returns a string value containing the short form of the path and name of
'*               the current file.
'*****************************************************************************************
Public Property Get ShortFileName() As String
    On Error GoTo hComponentFailure
    
    Dim sRet As String
    Dim iLen As Long
    
    ShortFileName = ""
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    iLen = Len(m_sPathFileName) + 1
    sRet = String$(iLen, 0)
    
    GetShortPathNameA m_sPathFileName, sRet, iLen
    ShortFileName = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    
    ShortFileName = sRet

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : Size
'* Notes       : Returns a long integer containing the size in bytes of the current file.
'*****************************************************************************************
Public Property Get Size() As Long
    On Error GoTo hComponentFailure
    
    Size = 0
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    Size = FileLen(m_sPathFileName)

    Exit Property

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Sub         : Copy
'* Notes       : Copies the current specified file (the source file) to a new file.
'*****************************************************************************************
Public Sub Copy(NewFileName As String)
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Len(NewFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_InvalidFileName, App.EXEName & ".CFile", S_ERR_InvalidFileName
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    lRet = CopyFile(Trim$(m_sPathFileName), Trim$(NewFileName), False)
    If lRet = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_CopyFailed, App.EXEName & ".CFile", S_ERR_CopyFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : Delete
'* Notes       : Deletes the current file.
'*****************************************************************************************
Public Sub Delete()
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
        
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
        
    lRet = DeleteFile(Trim$(m_sPathFileName))
    If lRet = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_DeleteFailed, App.EXEName & ".CFile", S_ERR_DeleteFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Function    : Exists
'* Notes       : Returns true if the specified file exist, false otherwise.
'*****************************************************************************************
Public Function Exists() As Boolean
    On Error GoTo hComponentFailure
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
    
    Exists = FileExists(m_sPathFileName)
    
    Exit Function
hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : IdenticalWith
'* Notes       : Returns True if the current file is identical (binary compare) with the
'*               specified file, False otherwise.
'*****************************************************************************************
Public Function IdenticalWith(ReferenceFileName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim fIsIdentical    As Boolean
    Dim iFF1            As Long
    Dim iFF2            As Long
    Dim lChunks         As Long
    Dim lFileLength     As Long
    Dim lFilePos        As Long
    Dim lFragment       As Long
    Dim X               As Long
    Dim buffer1         As String
    Dim buffer2         As String
    
    IdenticalWith = False
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Len(ReferenceFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_InvalidFileName, App.EXEName & ".CFile", S_ERR_InvalidFileName
    End If
    
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
        
    If Not FileExists(ReferenceFileName) Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
        
    If LCase$(m_sPathFileName) = LCase$(ReferenceFileName) Then
        IdenticalWith = True
    Else
        iFF1 = FreeFile
        Open m_sPathFileName For Binary As iFF1
        
        iFF2 = FreeFile
        Open ReferenceFileName For Binary As iFF2
        
        If LOF(iFF1) <> LOF(iFF2) Then
            IdenticalWith = False
        Else
            IdenticalWith = True
            
            lFileLength = LOF(iFF1)
            lChunks = lFileLength \ CHUNKSIZE
            lFragment = lFileLength Mod CHUNKSIZE
            
            buffer1$ = String$(CHUNKSIZE, 0)
            buffer2$ = String$(CHUNKSIZE, 0)
            
            lFilePos = 1
            For X = 1 To lChunks
                
                Get iFF1, lFilePos, buffer1
                Get iFF2, lFilePos, buffer2
                
                If StrComp(buffer1$, buffer2$, vbBinaryCompare) <> 0 Then
                    IdenticalWith = False
                    Exit For
                End If
                
                lFilePos = lFilePos + CHUNKSIZE
            Next
            
            buffer1 = String$(lFragment, 0)
            buffer2 = String$(lFragment, 0)
            
            Get iFF1, lFilePos, buffer1
            Get iFF2, lFilePos, buffer2
            
            If buffer1 <> buffer2 Then IdenticalWith = False
        End If
        
        Close iFF1
        Close iFF2
        
        IdenticalWith = fIsIdentical
    End If
    
    Exit Function

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : Move
'* Notes       : Moves the current specified file (the source file) to a new file.
'*****************************************************************************************
Public Sub Move(NewFileName As String)
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    
    If Len(m_sPathFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_NotInitialized, App.EXEName & ".CFile", S_ERR_NotInitialized
    End If
        
    If Len(NewFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_InvalidFileName, App.EXEName & ".CFile", S_ERR_InvalidFileName
    End If
        
    If Not Exists() Then
        On Error GoTo 0
        Err.Raise eErrFile_FileNotFound, App.EXEName & ".CFile", S_ERR_FileNotFound
    End If
    
    lRet = MoveFile(Trim$(m_sPathFileName), Trim$(NewFileName))
    If lRet = 0 Then
        On Error GoTo 0
        Err.Raise eErrFile_MoveFailed, App.EXEName & ".CFile", S_ERR_MoveFailed
    End If
    
    Exit Sub

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Function    : FileExists
'* Notes       : API-based routine to check for a file's existence.
'*               Returns True if the specified file exists, False otherwise.
'*****************************************************************************************
Private Function FileExists(FileName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim wfd As WIN32_FIND_DATA
    Dim hFile As Long
    
    FileExists = False
    
    hFile = FindFirstFile(FileName, wfd)
    If hFile <> INVALID_HANDLE_VALUE Then
        FileExists = True
        
        If FindClose(hFile) = 0 Then
            On Error GoTo 0
            Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
        End If
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrFile_ComponentFailure, App.EXEName & ".CFile", S_ERR_ComponentFailure
End Function

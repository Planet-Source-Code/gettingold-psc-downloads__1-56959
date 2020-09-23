VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcessFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'local variable(s) to hold property value(s)
Private mvarPrgID       As Long
Private mvarPrgTitle    As String                          'local copy
Private mvarPrgDesc     As String                          'local copy
Private mvarPrgURL      As String                          'local copy
Private mvarFileName    As String                          'local copy
Private mvarFileDate    As Date                            'local copy

Public Property Let PrgID(ByVal vdata As Long)
    mvarPrgID = vdata
End Property

Public Property Get PrgID() As Long
    PrgID = mvarPrgID
End Property

Public Property Let FileDate(ByVal vdata As Date)
    mvarFileDate = vdata
End Property

Public Property Get FileDate() As Date
    FileDate = mvarFileDate
End Property

Public Property Let FileName(ByVal vdata As String)
    mvarFileName = vdata
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let PrgURL(ByVal vdata As String)
    mvarPrgURL = vdata
End Property

Public Property Get PrgURL() As String
    PrgURL = mvarPrgURL
End Property

Public Property Let PrgDesc(ByVal vdata As String)
    mvarPrgDesc = vdata
End Property

Public Property Get PrgDesc() As String
    PrgDesc = mvarPrgDesc
End Property

Public Property Let PrgTitle(ByVal vdata As String)
    mvarPrgTitle = vdata
End Property

Public Property Get PrgTitle() As String
    PrgTitle = mvarPrgTitle
End Property

Public Sub DeleteTextFiles(strFileWithPath As String)

    Kill strFileWithPath

End Sub


Public Function ProcessTextFile(strFile As String) As Boolean

    Dim intFileHandle   As Integer
    Dim intPosn         As Integer
    Dim strOutput       As String

    intFileHandle = FreeFile

    'Open File
    Open strFile For Input As intFileHandle

    'Read 1st Line
    Line Input #intFileHandle, strOutput

    'Process Title
    PrgTitle = Trim(Mid(strOutput, 7))

    'Process the rest of the file
    Do Until EOF(intFileHandle)

        Line Input #intFileHandle, strOutput

        intPosn = InStr(1, strOutput, "Description:")
        If intPosn <> 0 Then
            PrgDesc = Trim(Mid(strOutput, 13))
        End If

        intPosn = InStr(1, strOutput, "vote on it at:")
        If intPosn <> 0 Then
            PrgURL = Trim(Mid(strOutput, intPosn + 14))
        End If

    Loop

    Close intFileHandle

    Call WriteRS

End Function

Public Function FileExist(strFile As String) As Boolean

    If PathFileExists(strFile) = 1 Then
        FileExist = True
    ElseIf PathFileExists(strFile) = 0 Then
        FileExist = False
    End If

End Function

Public Function WriteRS() As Boolean
    On Error GoTo WriteRS_ErrHandler

    With g_objRs
        .AddNew
        !FileName = FileName
        !FileDate = FileDate
        !PrgTitle = PrgTitle
        !PrgDesc = PrgDesc
        !PrgURL = PrgURL
        .Update
    End With

Proc_Exit:
    Exit Function

WriteRS_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in CProcessFiles: Function WriteRS"
    Resume Proc_Exit

End Function

Public Function Open_tblDLFiles() As Boolean
    On Error GoTo Open_tblDLFiles_ErrHandler

    Set g_objRs = New ADODB.Recordset

    g_objRs.Open "tbl_DLFiles", g_objCn, adOpenKeyset, adLockOptimistic

    Open_tblDLFiles = True

Proc_Exit:
    Open_tblDLFiles = False
    Exit Function

Open_tblDLFiles_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in CProcessFiles: Function Open_tblDLFiles"
    Resume Proc_Exit

End Function

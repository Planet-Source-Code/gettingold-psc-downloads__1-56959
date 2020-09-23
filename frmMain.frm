VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000002&
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   8400
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8115
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13864
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "10/27/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "9:30 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Default         =   -1  'True
      Height          =   315
      Left            =   5160
      Picture         =   "frmMain.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   315
   End
   Begin VB.ListBox lstDisp 
      BackColor       =   &H80000018&
      Height          =   7080
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Double-Click to see row details."
      Top             =   900
      Width           =   10815
   End
   Begin VB.ComboBox cboSearchIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "frmMain.frx":06D4
      Left            =   3240
      List            =   "frmMain.frx":06E4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   460
      Width           =   1875
   End
   Begin VB.TextBox txtSearchValue 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   460
      Width           =   3135
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Search In..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   2235
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search For..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_SearchForFile 
         Caption         =   "Search for File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile_SearchFolders 
         Caption         =   "Search Folder(s)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFile_Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oFF  As cFileFinder
Attribute oFF.VB_VarHelpID = -1
Private objZip          As XZip.Zip
Private objItem         As XZip.Item
Private strTempPath     As String

Private Sub SearchForFile(strFile As String)
    On Error GoTo SearchForFile_ErrHandler
    
    Dim objZip          As New XZip.Zip
    Dim clsProcessFiles As New CProcessFiles

    Screen.MousePointer = vbHourglass
    
    clsProcessFiles.Open_tblDLFiles
    
    For Each objItem In objZip.Contents(strFile)
        If Left(objItem.Name, 11) = "@PSC_ReadMe" Then
            Call objZip.UnPack(strFile, strTempPath, objItem.Name)
            clsProcessFiles.FileName = strFile
            clsProcessFiles.FileDate = FileDateTime(strTempPath & objItem.Name)
            Call clsProcessFiles.ProcessTextFile(strTempPath & objItem.Name)
            Call clsProcessFiles.DeleteTextFiles(strTempPath & objItem.Name)
            Exit For
        End If
    Next
        
    Screen.MousePointer = vbDefault

Proc_Exit:
    Call CloseRecordset(g_objRs)
    Set objZip = Nothing
    Set objItem = Nothing
    Set clsProcessFiles = Nothing
    Call Load_AllToListbox
    Exit Sub

SearchForFile_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in frmMain: Sub SearchForFile"
    Resume Proc_Exit

End Sub

Private Sub SearchForFiles(strSrchPath As String)

    Screen.MousePointer = vbHourglass

    oFF.Init "*.zip", strSrchPath, True, eFileFinderOutputStyle_FillFilesCollection
    oFF.FindFiles

    LoopThroughFoundFiles

    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSearch_Click()

    Call Load_AllToListbox
    
End Sub

Private Sub Form_Load()

    Set oFF = New cFileFinder
    
    'Get the systems temp path
    strTempPath = GetSystemTempPath
    
    ReDim TabArray(0 To 1) As Long
    
    TabArray(0) = 0
    TabArray(1) = 50
    
    'clear any existing tabs
    Call SendMessage(lstDisp.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    
    'set the list tabstops
    Call SendMessage(lstDisp.hwnd, LB_SETTABSTOPS, 2&, TabArray(0))
    lstDisp.Refresh
    
    'Load the data into the listbox
    Call Load_AllToListbox
    
    'Setup the progressbar
    prgStatus.Visible = False
    prgStatus.Move staStatus.Panels(1).Left + 40, staStatus.Top + 60, staStatus.Panels(1).Width - 100, staStatus.Height - 90
    
    'Set No Focus to command button
    Call NoFocusRect(cmdSearch, True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If g_objCn.State = adStateOpen Then
        g_objCn.Close
    End If
    
    Set g_objCn = Nothing
    Set oFF = Nothing

End Sub

Private Sub LoopThroughFoundFiles()
    On Error GoTo LoopThroughFoundFiles_ErrHandler

    Dim objZip          As New XZip.Zip
    Dim clsProcessFiles As New CProcessFiles
    Dim Count           As Long
    Dim iC              As Long
    Dim strFilName      As String

    Count = oFF.Count
    
    With Me.prgStatus
        .Min = 0
        .Max = Count
        .Visible = True
    End With
    
    clsProcessFiles.Open_tblDLFiles
    
    For iC = 1 To Count
        'Increment the progressbar
        prgStatus.Value = iC
        'Unzip all text files
        For Each objItem In objZip.Contents(oFF.Files(iC))
            If Left(objItem.Name, 11) = "@PSC_ReadMe" Then
                strFilName = oFF.StartPath & objItem.Name
                Call objZip.UnPack(oFF.Files(iC), strTempPath, objItem.Name)
                clsProcessFiles.FileName = oFF.Files(iC)
                clsProcessFiles.FileDate = FileDateTime(strTempPath & objItem.Name)
                Call clsProcessFiles.ProcessTextFile(strTempPath & objItem.Name)
                Call clsProcessFiles.DeleteTextFiles(strTempPath & objItem.Name)
                Exit For
            End If
        Next
    Next

Proc_Exit:
    Call CloseRecordset(g_objRs)
    Set objZip = Nothing
    Set objItem = Nothing
    Set clsProcessFiles = Nothing
    prgStatus.Visible = False
    Call Load_AllToListbox
    Exit Sub

LoopThroughFoundFiles_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in frmMain: Sub LoopThroughFoundFiles"
    Resume Proc_Exit

End Sub

Private Sub lstDisp_DblClick()
    
    frmDetails.txtPrgID = lstDisp.ItemData(lstDisp.ListIndex)
    
    frmDetails.Show vbModal
    
End Sub

Private Sub mnuFile_Exit_Click()

    Unload Me
    
End Sub

Private Function GetFolder() As String

    ' Displays the BrowseForFolder dialog and gives you the possibility to select a folder
    Dim oBrowse As CBrowseForFolder
    
    ' Basic object initialization
    Set oBrowse = New CBrowseForFolder
    oBrowse.Init "C:\", Me.hwnd, "Select a folder to Search."
    
    ' Show the dialog box
    If oBrowse.Browse Then
        Call SearchForFiles(oBrowse.SelectedPath)
    Else
        GetFolder = vbNullString
    End If
    
    ' Final cleanup
    Set oBrowse = Nothing

End Function

Private Function GetFile() As String

    ' Displays the standard File Open/Save dialog box
    Dim oFileDlg    As CFileOpenSaveDialog
    Dim iC          As Integer
    
    ' Basic object initialization
    Set oFileDlg = New CFileOpenSaveDialog
  
    ' Show the dialog box
    With oFileDlg
        .CenterDialog = vbChecked
        .DialogTitle = "Select the File to Add"
        .Filter = "All Files (*.*)|*.*|Zip Files (*.Zip)|*.Zip"
        .Flags = eFileOpenSaveFlag_Explorer + eFileOpenSaveFlag_FileMustExist + eFileOpenSaveFlag_HideReadOnly
        .HWndOwner = Me.hwnd
        .MaxFileSize = 255
        If .Show(eDialogType_OpenFile) Then
'            GetFile = .GetNextFileName
            Call SearchForFile(.GetNextFileName)
        End If
    End With

    ' Final cleanup
    Set oFileDlg = Nothing

End Function

Private Sub mnuFile_SearchFolders_Click()

    Call GetFolder
    
End Sub

Private Sub mnuFile_SearchForFile_Click()

    Call GetFile
    
End Sub

Private Sub Load_AllToListbox()
    On Error GoTo Load_AllToListbox_ErrHandler
    
    Dim strSQL      As String
    Dim strFilter   As String
    
    Call LockWindowUpdate(Me.hwnd)
    
    lstDisp.Clear
    
    Set g_objRs = New ADODB.Recordset
    
    g_objRs.Open "qry_SortedTable", g_objCn, adOpenForwardOnly, adLockReadOnly
    
    If txtSearchValue <> vbNullString Then
        If cboSearchIn.ListIndex >= 0 Then
            Select Case cboSearchIn.ListIndex
                Case 0 'File Name
                    g_objRs.Filter = "FileName Like '*" & txtSearchValue & "*'"
                Case 1 'Title
                    g_objRs.Filter = "prgTitle Like '*" & txtSearchValue & "*'"
                Case 2 'Description
                    g_objRs.Filter = "prgDesc Like '*" & txtSearchValue & "*'"
                Case 3 'Title & Description
                    g_objRs.Filter = "prgTitle Like '*" & txtSearchValue & "*' Or prgDesc Like '*" & txtSearchValue & "*'"
                Case Else
            End Select
        Else
            MsgBox "Select Fields to Search in.", vbInformation
        End If
    End If
    
    If RecordsFound(g_objRs) = True Then
        Do Until g_objRs.EOF
            Call AddToListbox(lstDisp, Format(g_objRs!FileDate, "mm/dd/yyyy") & vbTab & g_objRs!PrgTitle, g_objRs!PrgID)
            g_objRs.MoveNext
        Loop
    End If
    
Proc_Exit:
    Call CloseRecordset(g_objRs)
    staStatus.Panels(1).Text = lstDisp.ListCount & " Files Found"
    Call LockWindowUpdate(0)
    Exit Sub

Load_AllToListbox_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in frmMain: Sub Load_AllToListbox"
    Resume Proc_Exit

End Sub

Private Sub txtSearchValue_GotFocus()

    Call AutoSelect
    
End Sub

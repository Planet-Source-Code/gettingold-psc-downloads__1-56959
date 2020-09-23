Attribute VB_Name = "modCommon"
Option Explicit

Private StandardButtonProc  As Long

Public g_objCn  As Connection
Public g_objRs  As Recordset

Private Const GW_HWNDPREV = 3
Private Const LB_ADDSTRING = &H180
Private Const LB_SETITEMDATA = &H19A
Private Const GWL_WNDPROC = (-4)
Private Const WM_SETFOCUS = &H7
Private Const MAX_PATH = 260

Public Const LB_SETTABSTOPS As Long = &H192

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub ActivatePrevInstance()

    Dim OldTitle   As String
    Dim PrevHndl   As Long
    Dim result     As Long
    
    'Save the title of the application.
    OldTitle = App.Title
    
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    App.Title = "unwanted instance"
     
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    
    'Check if found
    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    End If
    
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    
    'Restore the program.
    result = OpenIcon(PrevHndl)
    
    'Activate the application.
    result = SetForegroundWindow(PrevHndl)
    
    'End the application.
    End
    
End Sub

Public Sub CloseRecordset(ByRef rs As ADODB.Recordset)

    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Set rs = Nothing
    
End Sub

Public Function RecordsFound(ByRef rs As ADODB.Recordset) As Boolean

    If Not rs.BOF And Not rs.EOF Then
        RecordsFound = True
    Else
        RecordsFound = False
    End If
    
End Function

Private Function OpenTheDatabase() As Boolean
    On Error GoTo OpenTheDatabase_ErrHandler

    Dim strDbPath   As String
    
    strDbPath = App.Path & "\DLF.mdb"
    
    Set g_objCn = New ADODB.Connection
    
    g_objCn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & ";Persist Security Info=False" ';Jet OLEDB:Database Password=~M28d*5J%cb27"
    g_objCn.Open

    OpenTheDatabase = True

Proc_Exit:
    Exit Function

OpenTheDatabase_ErrHandler:
    OpenTheDatabase = False
    MsgBox "Database was not opened, program terminated." & vbNewLine & vbNewLine & Err.Description, , "Error #" & Err.Number
    Resume Proc_Exit

End Function

Public Sub Main()

    Call ActivatePrevInstance
    If OpenTheDatabase = True Then
        frmMain.Show
    End If
    
End Sub

'-------------------------------------------------------------------------------------
' Sub AddToListbox
'
' Created by Mark Bader
' Date: 01-29-2002
'
' Purpose: This routine add items & ItemData to a ComboBox using API's
'
'    objList:       ListBox to be loaded
'    bClearLst:     Clear the Listbox if True
'    strAddItem:    String to be loaded into listbox
'    lngItemData:   Long Number to be added into ItemData Field of ListBox (Optional)
'
'-------------------------------------------------------------------------------------
Public Sub AddToListbox(objListbox As ListBox, strAddItem As String, Optional lngItemData As Long)

    Dim lngIndex    As Long

    lngIndex = SendMessage(objListbox.hwnd, LB_ADDSTRING, 0, ByVal strAddItem)

    If Not IsMissing(lngItemData) Then
        Call SendMessage(objListbox.hwnd, LB_SETITEMDATA, lngIndex, ByVal lngItemData)
    End If

End Sub

Public Sub NoFocusRect(Button As Object, vValue As Boolean)

    If vValue = True Then 'Focus rect on
        'Save the adress of the standard button procedure
        StandardButtonProc = GetWindowLong(Button.hwnd, GWL_WNDPROC)
        'Subclass the button to control its Windows Messages
        SetWindowLong Button.hwnd, GWL_WNDPROC, AddressOf ButtonProc
    Else 'Focus rect off
        'Remove the subclassing from the button
        SetWindowLong Button.hwnd, GWL_WNDPROC, StandardButtonProc
    End If
    
End Sub

Private Function ButtonProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next

    'The procedure that gets all windows messages for the subclassed button
    
    Select Case uMsg&
        'The button is going to get the focus
        Case WM_SETFOCUS
        'Exit the procedure -> The message doesn´t reach the button
        Exit Function
    End Select
    
    'Call the standard Button Procedure
    ButtonProc = CallWindowProc(StandardButtonProc, hwnd&, uMsg&, wParam&, lParam&)
    
End Function

Sub AutoSelect()
    On Error Resume Next
    
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)

End Sub

Public Function GetSystemTempPath() As String
   
   Dim result As Long
   Dim buff As String
   
  'get the user's \temp folder
  'pad the passed string
   buff = Space$(MAX_PATH)
   
   result = GetTempPath(MAX_PATH, buff)
   
  'result contains the number of chrs up to the
  'terminating null, so a simple left$ can
  'be used. Its also conveniently terminated
  'with a slash.
   GetSystemTempPath = Left$(buff, result)

End Function


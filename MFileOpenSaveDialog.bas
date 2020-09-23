Attribute VB_Name = "MFileOpenSaveDialog"
Option Explicit

' Private type definitions
Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type

' Private constants
Private Const WM_INITDIALOG = &H110

' Private API function declarations
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long


'*****************************************************************************************
'* Function    : FileOpenSaveDialogCallbackEx
'* Notes       : Old-style callback (global) routine for CFileOpenSaveDialog. It is used
'*               to center the dialog box. Must be used with eFileOpenSaveFlag_Explorer.
'*****************************************************************************************
Public Function FileOpenSaveDialogCallbackEx(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight     As Long
    Dim rcWidth      As Long
    Dim HWndParent   As Long
    Dim rc           As RECT
    Dim rcDesk       As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            ' Center the window
            HWndParent = GetParent(hwnd)

            Call GetWindowRect(GetDesktopWindow, rcDesk)
            Call GetWindowRect(HWndParent, rc)

            rcHeight = rc.Bottom - rc.Top
            rcWidth = rc.Right - rc.Left
            rc.Left = (((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
            rc.Top = (((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

            MoveWindow HWndParent, rc.Left, rc.Top, rcWidth, rcHeight, 1

        Case Else
            '

    End Select

    FileOpenSaveDialogCallbackEx = 0&
End Function


'*****************************************************************************************
'* Function    : FileOpenSaveDialogCallback
'* Notes       : Old-style callback (global) routine for CFileOpenSaveDialog. It is used
'*               to center the dialog box. Does not cover eFileOpenSaveFlag_Explorer.
'*****************************************************************************************
Public Function FileOpenSaveDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight     As Long
    Dim rcWidth      As Long
    Dim rc           As RECT
    Dim rcDesk       As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            ' Center the window
            Call GetWindowRect(GetDesktopWindow, rcDesk)
            Call GetWindowRect(hwnd, rc)

            rcHeight = rc.Bottom - rc.Top
            rcWidth = rc.Right - rc.Left
            rc.Left = (((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
            rc.Top = (((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

            MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1

        Case Else
            '

    End Select

    FileOpenSaveDialogCallback = 0&
End Function

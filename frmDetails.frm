VERSION 5.00
Begin VB.Form frmDetails 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Details"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtprgURL 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   3525
      Width           =   8715
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Close"
      Height          =   330
      Index           =   1
      Left            =   7560
      TabIndex        =   11
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Frame fra 
      Height          =   60
      Left            =   60
      TabIndex        =   9
      Top             =   4500
      Width           =   8835
   End
   Begin VB.TextBox txtPrgID 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   2460
      TabIndex        =   8
      Top             =   4125
      Width           =   6375
   End
   Begin VB.TextBox txtFileDate 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   4125
      Width           =   2415
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H80000018&
      Height          =   1995
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1245
      Width           =   8715
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   405
      Width           =   8715
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   " PSC URL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   8715
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   "File Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   6315
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   " File Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   8715
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Program Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8715
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsProcessFiles As CProcessFiles

Private Sub cmdAction_Click(Index As Integer)

    If Index = 1 Then Unload Me
    
End Sub

Private Sub Form_Activate()

    'create object
    Set clsProcessFiles = New CProcessFiles
    
    'Get the record
    Call Get_the_Record

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set clsProcessFiles = Nothing
    
End Sub

Private Sub Get_the_Record()
    On Error GoTo Get_the_Record_ErrHandler

    Dim strSQL  As String
    
    Set g_objRs = New ADODB.Recordset
    
    strSQL = "Select * From tbl_DLFiles Where prgID = " & txtPrgID

    g_objRs.Open strSQL, g_objCn, adOpenForwardOnly, adLockReadOnly
    
    If RecordsFound(g_objRs) = True Then
        txtDesc = g_objRs!PrgDesc
        txtFileDate = g_objRs!FileDate
        txtFileName = g_objRs!FileName
        txtTitle = g_objRs!PrgTitle
        txtprgURL = g_objRs!PrgURL
    End If

Proc_Exit:
    Call CloseRecordset(g_objRs)
    Exit Sub

Get_the_Record_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in frmDetails: Sub Get_the_Record"
    Resume Proc_Exit

End Sub

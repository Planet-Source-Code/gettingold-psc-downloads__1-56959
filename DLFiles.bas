Attribute VB_Name = "DLFiles"
Option Explicit

Public Function Exec_qry_del_tbl_DLFiles(ByVal lngPrgID As Long) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

        On Error GoTo PROC_ERR

        strSQL = "qry_del_tbl_DLFiles"
        With objCmd
                .CommandText = strSQL
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = g_objCn

                .Parameters.Append .CreateParameter("pPrgID", adInteger, adParamInput, 4, lngPrgID)
        
                .Execute Options:=adExecuteNoRecords
        End With

        Set objCmd = Nothing

        Exec_qry_del_tbl_DLFiles = 0
        Exit Function
PROC_ERR:
        Exec_qry_del_tbl_DLFiles = Err.Number
End Function

Public Function Exec_qry_sel_tbl_DLFiles(ByVal lngPrgID As Long, ByRef objRs As ADODB.Recordset) As Long
    On Error GoTo Exec_qry_sel_tbl_DLFiles_ErrHandler

    Dim strSQL  As String
    Dim objCmd  As New ADODB.Command

    strSQL = "qry_sel_tbl_DLFiles"
    
    With objCmd
    
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        
        Set .ActiveConnection = g_objCn

        .Parameters.Append .CreateParameter("pPrgID", adInteger, adParamInput, 4, lngPrgID)
        
        objRs.Open objCmd
        
    End With

    Set objCmd = Nothing

    Exec_qry_sel_tbl_DLFiles = 0

Proc_Exit:
    Exit Function

Exec_qry_sel_tbl_DLFiles_ErrHandler:
    MsgBox Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")", , "Error in DLFiles: Function Exec_qry_sel_tbl_DLFiles"
    Resume Proc_Exit

End Function

Public Function Exec_qry_ins_tbl_DLFiles(ByVal strPrgTitle As String, ByVal strPrgDesc As String, ByVal strPrgURL As String, ByVal strFileName As String, ByVal dteFileDate As Date) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

        On Error GoTo PROC_ERR

        strSQL = "qry_ins_tbl_DLFiles"
        With objCmd
                .CommandText = strSQL
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = g_objCn

                .Parameters.Append .CreateParameter("pPrgTitle", adVarWChar, adParamInput, 225, strPrgTitle)
                .Parameters.Append .CreateParameter("pPrgDesc", adLongVarWChar, adParamInput, 2147483647, strPrgDesc)
                .Parameters.Append .CreateParameter("pPrgURL", adVarWChar, adParamInput, 225, strPrgURL)
                .Parameters.Append .CreateParameter("pFileName", adVarWChar, adParamInput, 225, strFileName)
                .Parameters.Append .CreateParameter("pFileDate", adDBTimeStamp, adParamInput, 8, dteFileDate)
        
                .Execute Options:=adExecuteNoRecords
        End With

        Set objCmd = Nothing

        Exec_qry_ins_tbl_DLFiles = 0
        Exit Function
PROC_ERR:
        Exec_qry_ins_tbl_DLFiles = Err.Number
End Function

Public Function Exec_qry_upd_tbl_DLFiles(ByVal lngPrgID As Long, ByVal strPrgTitle As String, ByVal strPrgDesc As String, ByVal strPrgURL As String, ByVal strFileName As String, ByVal dteFileDate As Date) As Long
 Dim strSQL As String
 Dim objCmd As New ADODB.Command

        On Error GoTo PROC_ERR

        strSQL = "qry_upd_tbl_DLFiles"
        With objCmd
                .CommandText = strSQL
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = g_objCn

                .Parameters.Append .CreateParameter("pPrgTitle", adVarWChar, adParamInput, 225, strPrgTitle)
                .Parameters.Append .CreateParameter("pPrgDesc", adLongVarWChar, adParamInput, 2147483647, strPrgDesc)
                .Parameters.Append .CreateParameter("pPrgURL", adVarWChar, adParamInput, 225, strPrgURL)
                .Parameters.Append .CreateParameter("pFileName", adVarWChar, adParamInput, 225, strFileName)
                .Parameters.Append .CreateParameter("pFileDate", adDBTimeStamp, adParamInput, 8, dteFileDate)
                .Parameters.Append .CreateParameter("pPrgID", adInteger, adParamInput, 4, lngPrgID)
        
                .Execute Options:=adExecuteNoRecords
        End With

        Set objCmd = Nothing

        Exec_qry_upd_tbl_DLFiles = 0
        Exit Function
PROC_ERR:
        Exec_qry_upd_tbl_DLFiles = Err.Number
End Function

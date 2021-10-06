Attribute VB_Name = "mDemo"
Option Explicit

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Public Sub ErrorHandling_None_Demo()
    Dim l As Long
    l = ErrorHandling_None(10, 0)
End Sub

Private Function ErrorHandling_None(ByVal op1 As Variant, _
                                    ByVal op2 As Variant) As Variant
' ------------------------------------------------------------------
' - Error message:              Mere VBA only
'   - Error source:             No
'   - Application error number: Not supported
'   - Error line:               No, even when one is provided/available
'   - Info about error:         Not supported
'   - Path to the error:        No, because a call stack is not maintained
' - Variant value assertion:    No
' - Execution Trace:            No
' - Debugging/Test choice:      No
' ------------------------------------------------------------------
    ErrorHandling_None = op1 / op2
End Function

Public Sub Demo_6_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Demo_6_Execution_Trace"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6a
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_no:=Err.Number, err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description, err_line:=Erl
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6a()

    Const PROC = "Demo_6_Execution_Trace_DemoProc_6a"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6b

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_no:=Err.Number, err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description, err_line:=Erl
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6b()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6b"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    
    Demo_6_Execution_Trace_DemoProc_6c
    
    Dim i As Long: Dim j As Long: j = 10000000
    mTrc.BoC ErrSrc(PROC) & " empty loop 1 to " & j
    For i = 1 To j
    Next i
    mTrc.EoC ErrSrc(PROC) & " empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_no:=Err.Number, err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description, err_line:=Erl
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6c()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_no:=Err.Number, err_source:=ErrSrc(PROC), err_dscrptn:=Err.Description, err_line:=Erl
End Sub

Private Function ErrSrc(ByVal s As String) As String
' ---------------------------------------------------
' Prefix procedure name (s) by this module's name.
' Attention: The characters > and < must not be used!
' ---------------------------------------------------
    ErrSrc = "mDemo." & s
End Function

Private Sub ErrMsgMatter(ByVal err_source As String, _
                         ByVal err_no As Long, _
                         ByVal err_line As Long, _
                         ByVal err_dscrptn As String, _
                Optional ByRef msg_title As String, _
                Optional ByRef msg_type As String, _
                Optional ByRef msg_line As String, _
                Optional ByRef msg_no As Long, _
                Optional ByRef msg_details As String, _
                Optional ByRef msg_dscrptn As String, _
                Optional ByRef msg_info As String)
' -------------------------------------------------------
' Returns all the matter to build a proper error message.
' -------------------------------------------------------
                
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)     ' Message error line
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)                ' Message error number
    msg_title = msg_type & msg_no & " in " & err_source & " " & msg_line             ' Message title
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in " & err_source & " (at line " & err_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1)

End Sub

Private Sub ErrMsg(ByVal err_no As Long, _
                   ByVal err_source As String, _
                   ByVal err_dscrptn As String, _
                   ByVal err_line As Long)
' ----------------------------------------------
'
' ----------------------------------------------
    Dim sTitle As String
    
    ErrMsgMatter err_source:=err_source, err_no:=err_no, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_title:=sTitle
    
    MsgBox Prompt:="Error description:" & vbLf & _
                    err_dscrptn & vbLf & vbLf & _
                   "Error source:" & vbLf & _
                   err_source, _
           Buttons:=vbOKOnly, _
           Title:=sTitle
    mTrc.Finish sTitle
    mTrc.Terminate
End Sub



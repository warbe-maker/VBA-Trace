Attribute VB_Name = "mTest"
Option Explicit

Public Const CONCAT = "||"
' ----------------------------------------------------------------------
' Deklarations for the use of the fMsg UserForm (Alternative VBA MsgBox)
Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------
' ----------------------------------------------------------------------

Private bRegressionTest As Boolean

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

Private Function RegressionTestInfo() As String
' ----------------------------------------------------
' Adds s to the Err.Description as an additional info.
' ----------------------------------------------------
    RegressionTestInfo = err.Description
    If Not bRegressionTest Then Exit Function
    
    If InStr(RegressionTestInfo, CONCAT) <> 0 _
    Then RegressionTestInfo = RegressionTestInfo & vbLf & vbLf & "Please notice that  this is a  r e g r e s s i o n  t e s t ! Click any but the ""Terminate"" button to continue with the test in case another one follows." _
    Else RegressionTestInfo = RegressionTestInfo & CONCAT & "Please notice that  this is a  r e g r e s s i o n  t e s t !  Click any but the ""Terminate"" button to continue with the test in case another one follows."

End Function

Public Sub Regression_Test()
' -----------------------------------------------------------------------------
' 1. This regression test requires the Conditional Compile Argument "Test = 1"
'    which provides additional buttons to continue with the next test after a
'    procedure which tests an error condition
' 2. The BoP/EoP statements in this regression test procedure produce one
'    execution trace at the end of this regression test provided the
'    Conditional Compile Argument "ExecTrace = 1". Attention must be paid on
'    the execution time however because it will include the time used by the
'    user action when an error message is displayed!
' 3. The Conditional Compile Argument "Debugging = 1" allows to identify the
'    code line which causes the error through an extra "Resume error code line"
'    button displayed with the error message and processed when clicked as
'    "Stop: Resume" when the button is clicked.
' ------------------------------------------------------------------------------
    
    On Error GoTo eh
    Const PROC = "Regression_Test"
    bRegressionTest = True
    
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace
    Test_3_Execution_Trace_With_Error

xt: mTrc.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Public Sub Test_1_1_BoP_missing()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Test_1_1_BoP_missing"
    
'    mTrc.BoP ErrSrc(PROC) this procedure will not be recognized as "Entry Procedure" ...
    Test_1_1_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_1_1_BoP_missing_TestProc_1a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Test_1_1_BoP_missing_TestProc_1a"
    
    mTrc.BoP ErrSrc(PROC)
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub


Public Sub Test_1_2_BoP_missing()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Test_1_2_BoP_missing"
    
    mTrc.BoP ErrSrc(PROC)
    Test_1_2_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_1_2_BoP_missing_TestProc_1a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Test_1_2_BoP_missing_TestProc_1a"
    
'    mTrc.BoP ErrSrc(PROC)
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Public Sub Test_2_BoP_EoP()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Test_2_BoP_EoP"
    
    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1a_missing_BoP
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1a_missing_BoP()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Test_2_BoP_EoP_TestProc_1a_missing_BoP"
    
'    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1b_paired_BoP_EoP
    Test_2_BoP_EoP_TestProc_1d_missing_EoP

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1b_paired_BoP_EoP()
    Const PROC = "Test_2_BoP_EoP_TestProc_1b_paired_BoP_EoP"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1c_missing_EoC
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub
    
eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1c_missing_EoC()
    Const PROC = "Test_2_BoP_EoP_TestProc_1c_missing_EoC"
    
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    BoC ErrSrc(PROC) & " trace of some code lines" ' missing EoC statement

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub
    
eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1e_BoC_EoC()
    Const PROC = "Test_2_BoP_EoP_TestProc_1e_BoC_EoC"
    
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
        
    Dim i As Long: Dim j As Long: j = 10000000
    BoC PROC & " code trace empty loop 1 to " & j
    For i = 1 To j
    Next i
    EoC PROC & " code trace empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1d_missing_EoP()
    Const PROC = "Test_2_BoP_EoP_TestProc_1d_missing_EoP"
    
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1e_BoC_EoC
    
xt: Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Public Sub Test_3_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_3_Execution_Trace"
    On Error GoTo eh
'    mTrc.DisplayedInfo = Compact
    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace_TestProc_6a "arg1", "arg2", 15.2

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_3_Execution_Trace_TestProc_6a(ByVal arg1 As Variant, _
                                               ByVal arg2 As Variant, _
                                               ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Test_3_Execution_Trace_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC), arg1, arg2, arg3
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_3_Execution_Trace_TestProc_6b
    Test_3_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_3_Execution_Trace_TestProc_6b()
    
    Const PROC = "Test_3_Execution_Trace_TestProc_6b"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_3_Execution_Trace_TestProc_6c()
    
    Const PROC = "Test_3_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub


Public Sub Test_3_Execution_Trace_With_Error()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_3_Execution_Trace_With_Error"
    On Error GoTo eh
    mTrc.DisplayedInfo = Compact
'    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace_With_Error_TestProc_6a
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_3_Execution_Trace_With_Error_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Test_3_Execution_Trace_With_Error_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_3_Execution_Trace_With_Error_TestProc_6b
    Test_3_Execution_Trace_With_Error_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Sub Test_3_Execution_Trace_With_Error_TestProc_6b()
    
    Const PROC = "Test_3_Execution_Trace_With_Error_TestProc_6b"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC)
End Sub

Private Sub Test_3_Execution_Trace_With_Error_TestProc_6c()
    
    Const PROC = "Test_3_Execution_Trace_With_Error_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    Dim i As Long: i = i / 0 ' Error !!!!

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err_source:=ErrSrc(PROC), err_asserted:=6
End Sub

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
' ---------------------------------------------------------------------------------
' Returns all matter to build a proper error message.
' msg_line:    at line <err_line>
' msg_no:      1 to n (an Application error translated back into its origin number)
' msg_title:   <error type> <error number> in:  <error source>
' msg_details: <error type> <error number> in <error source> [(at line <err_line>)]
' msg_dscrptn: the error description
' msg_info:    any text which follows the description concatenated by a ||
' ---------------------------------------------------------------------------------
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)
    msg_title = msg_type & msg_no & " in:  " & err_source
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in: " & err_source & " (" & msg_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1) Else msg_info = vbNullString

End Sub

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString, _
    Optional ByVal err_line As Long = 0, _
    Optional ByVal err_asserted = 0)
' --------------------------------------------------
' Note! Because the mTrc trace module is an optional
'       module of the mErH error handler module it
'       cannot use the mErH's ErrMsg procedure and
'       thus uses its own - with the known
'       disadvantage that the title maybe truncated.
' --------------------------------------------------
    Dim sTitle      As String
    Dim sDetails    As String
    
    If err_no = 0 Then err_no = err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = err.Description
    If err_line = 0 Then err_line = Erl
    
    ErrMsgMatter err_source:=err_source, err_no:=err_no, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_title:=sTitle, msg_details:=sDetails
    
#If Test Then
    If err_no <> err_asserted _
    Then MsgBox Prompt:="Error description:" & vbLf & _
                        err_dscrptn & vbLf & vbLf & _
                        "Error source/details:" & vbLf & _
                        sDetails, _
                buttons:=vbOKOnly, _
                Title:=sTitle
#Else
    MsgBox Prompt:="Error description:" & vbLf & _
                    err_dscrptn & vbLf & vbLf & _
                   "Error source/details:" & vbLf & _
                   sDetails, _
           buttons:=vbOKOnly, _
           Title:=sTitle
#End If
    mTrc.Finish sTitle
    mTrc.Terminate
End Sub


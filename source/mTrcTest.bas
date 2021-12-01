Attribute VB_Name = "mTrcTest"
Option Explicit
' -----------------------------------------------------------------------
' Standar module mTest: Provides all test obligatory being executed when
'                       code in mTrc is modified.
'
' -----------------------------------------------------------------------
Public Const CONCAT = "||"

Private bRegressionTest As Boolean

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is a kind of universal error message which includes a debugging option.
' It may be copied into any module - turned into a Private function. When the/my
' Common VBA Error Handling Component (ErH) is installed and the Conditional
' Compile Argument 'CommErHComp = 1' the error message will be displayed by
' means of the Common VBA Message Component (fMsg, mMsg).
'
' Usage: When this procedure is copied as a Private Function into any desired
'        module an error handling which consideres the possible Conditional
'        Compile Argument 'Debugging = 1' will look as follows
'
'            Const PROC = "procedure-name"
'            On Error Goto eh
'        ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume
'               Case vbNo:  Resume Next
'               Case Else:  Goto xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Used:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              err_dscrptn & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
#If CommErHComp Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
#Else
    '~~ When the Common VBA Error Handling Component (ErH) is not used/installed there might still be the
    '~~ Common VBA Message Component (Msg) be installed/used
#If CommMsgComp Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
#Else
    '~~ None of the Common Components is installed/used
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
#End If
#End If
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

Private Function RegressionTestInfo() As String
' ----------------------------------------------------
' Adds s to the Err.Description as an additional info.
' ----------------------------------------------------
    RegressionTestInfo = Err.Description
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
    Const PROC = "Regression_Test"
    
    On Error GoTo eh
    bRegressionTest = True
    
    mTrc.DisplayedInfo = Compact
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace
    Test_3_Execution_Trace_With_Error

xt: mTrc.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_1_1_BoP_missing()
' ---------------------------------------------------
' About this test:
' Because BoP/EoP is inconsistent, the EoP is ignored and the called procedure
' which has consistent BoP/EoP statements is execution traced
' ---------------------------------------------------
    Const PROC = "Test_1_1_BoP_missing"
    
'    mTrc.BoP ErrSrc(PROC) this procedure will not be recognized as "Entry Procedure" ...
    Test_1_1_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1b_paired_BoP_EoP()
    Const PROC = "Test_2_BoP_EoP_TestProc_1b_paired_BoP_EoP"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1c_paired_BoP_EoP
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1c_paired_BoP_EoP()
    Const PROC = "Test_2_BoP_EoP_TestProc_1c_paired_BoP_EoP"
    
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    BoC ErrSrc(PROC) & " trace of some code lines (EoC statement missing!)"

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_2_BoP_EoP_TestProc_1d_missing_EoP()
    Const PROC = "Test_2_BoP_EoP_TestProc_1d_missing_EoP"
    
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Test_2_BoP_EoP_TestProc_1e_BoC_EoC
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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
    
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace_TestProc_6a arg1:="xxxx", arg2:="yyyy", arg3:=12.8

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_3_Execution_Trace_TestProc_6a(ByVal arg1 As Variant, _
                                               ByVal arg2 As Variant, _
                                               ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Test_3_Execution_Trace_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC), arg1, "arg2=", arg2, arg3
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_3_Execution_Trace_TestProc_6b
    Test_3_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_3_Execution_Trace_TestProc_6c()
    
    Const PROC = "Test_3_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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
    
    mTrc.BoP ErrSrc(PROC)
    Test_3_Execution_Trace_With_Error_TestProc_6a
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_3_Execution_Trace_With_Error_TestProc_6c()
    
    Const PROC = "Test_3_Execution_Trace_With_Error_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    Dim i As Long
    i = i / 0 ' Error !!!!

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Resume Next
        Case Else:  GoTo xt
    End Select
End Sub



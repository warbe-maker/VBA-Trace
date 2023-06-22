Attribute VB_Name = "TestmTrc"
Option Explicit
' -----------------------------------------------------------------------
' Standard Module TestmTrc: Provides all test obligatory being executed
' ========================= when the mTrc code is modified.
'
'
' Uses (for test only): mBasic, fMsg/mMsg, mErH, mTrc
'
' W. Rauschenberger Berlin June 2023
' -----------------------------------------------------------------------

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

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "TestmTrc." & s
End Function

Public Sub Regression_Test()
' ----------------------------------------------------------------------------
' Attention! Testing the clsmTrc module requires the Cond. Comp. Arg.:
' ========== XcTrc_clsTrc = 1 : XcTrc_mTrc = 0
'            Testing the mTrc module requires the Cond. Comp. Arg.:
'            XcTrc_clsTrc = 0: XcTrc_mTrc = 1
'            !! Performing both after a code change (usually consistent in
'            !! both componentents) is obligatory !
'
' The resulting trace-log-file's content is displayed by means of ShellRun.
' ------------------------------------------------------------------------------
    Const PROC = "Regression_Test"
    
    On Error GoTo eh
    
    '~~ Initializations (must be done prior the first BoP !)
    mTrc.FileName = "RegressionTest_mTrc.ExecTrace.log"
    mTrc.Title = "Regression Test Standard Module mTrc"
    mTrc.NewFile
    mErH.Regression = True
        
    mBasic.BoP ErrSrc(PROC), "arg1, arg2"
    Regression_Test_03_Execution_Trace
    mTrc.LogInfo = "Test Log-Info explicitly provided"
    Regression_Test_03_Execution_Trace_With_Error

xt: mBasic.EoP ErrSrc(PROC), "arg1, arg2"
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_Test_01_1_BoP_missing()
' ---------------------------------------------------
' About this test:
' Because BoP/EoP is inconsistent, the EoP is ignored and the called procedure
' which has consistent BoP/EoP statements is execution traced
' ---------------------------------------------------
    Const PROC = "Regression_Test_01_1_BoP_missing"
    
'    mBasic.BoP ErrSrc(PROC) this procedure will not be recognized as "Entry Procedure" ...
    Regression_Test_01_1_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_01_1_BoP_missing_TestProc_1a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_01_1_BoP_missing_TestProc_1a"
    
    mBasic.BoP ErrSrc(PROC)
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_Test_01_2_BoP_missing()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Regression_Test_01_2_BoP_missing"
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_01_2_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_01_2_BoP_missing_TestProc_1a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_01_2_BoP_missing_TestProc_1a"
    
'    mBasic.BoP ErrSrc(PROC)
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_Test_02_BoP_EoP()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Regression_Test_02_BoP_EoP"
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1a_missing_BoP
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1a_missing_BoP()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1a_missing_BoP"
    
'    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP
    Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP"
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mBasic.BoC ErrSrc(PROC) & " trace of some code lines (EoC statement missing!)"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
        
    Dim i As Long: Dim j As Long: j = 10000000
    mBasic.BoC PROC & " code trace empty loop 1 to " & j
    For i = 1 To j
    Next i
    mBasic.EoC PROC & " code trace empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_Test_03_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Regression_Test_03_Execution_Trace"
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_03_Execution_Trace_TestProc_6a arg1:="xxxx", arg2:="yyyy", arg3:=12.8

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_6a(ByVal arg1 As Variant, _
                                               ByVal arg2 As Variant, _
                                               ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6a"
    
    mBasic.BoP ErrSrc(PROC), arg1 & " arg2=" & arg2 & ", " & arg3
    mBasic.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_TestProc_6b
    Regression_Test_03_Execution_Trace_TestProc_6c
    mBasic.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_6b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6b"
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_6c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_Test_03_Execution_Trace_With_Error()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error"
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6a
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6a"
    
    mBasic.BoP ErrSrc(PROC)
    mBasic.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6b
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6c
    mBasic.EoC ErrSrc(PROC) & " call of 6b and 6c"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6b"
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6c"
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    '~~ The VB Runtime error 6 is anticipated thus regarded asserted
    '~~ when mErH.Regression = True for this test (set with the
    '~~ calling procedure) the display of the error is suspended
    mErH.Asserted 6 ' VB-Runtime-error overflow
    mTrc.LogInfo = "This procedure will raise a VB-Runtime Error 6 (oderflow) - not displayed with the regression test because asserted by mErH.Assert"
    Dim i As Long
    i = i / 0 ' Error !!!!

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_99_DefaultVersusSpecifiedLogFile()

    Dim fso             As New FileSystemObject
    Dim s               As String
    Dim sFileName       As String
    Dim sPath           As String
    Dim sFileFullName   As String
    
    sFileName = fso.GetBaseName(ThisWorkbook.Name) & ".ExecTrace.log"
    sPath = ThisWorkbook.Path
    sFileFullName = sPath & "\" & sFileName ' creates the file

    '~~ 1. Defaults
    mTrc.Initialize ' The way for the Standard Module to provide defaults
    Debug.Assert mTrc.FileFullName = sFileFullName
    mTrc.Path = sPath
    mTrc.FileName = sFileName
    Debug.Assert mTrc.FileFullName = sFileFullName
    
    '~~ 2. NewFile test
    mTrc.NewFile
    Debug.Assert fso.FileExists(mTrc.FileFullName)
    
    '~~ 3. Go with user-spec log-file (existing default is deleted)
    mTrc.Initialize ' setup defaults
    mTrc.FileName = "ExecTrace.My.log"
    Debug.Assert mTrc.FileFullName = sPath & "\ExecTrace.My.log"
    Debug.Assert Not fso.FileExists(mTrc.FileFullName)
        
    Set fso = Nothing
    
End Sub


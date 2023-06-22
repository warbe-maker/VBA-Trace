Attribute VB_Name = "TestclsTrc"
Option Explicit
' -----------------------------------------------------------------------
' Standard Module TestclsTrc: Provides all test obligatory being executed
' =========================== when the clsTrc code is modified.
'
' Uses (for test only): mBasic, fMsg/mMsg, mErH, clsTrc
'
' W. Rauschenberger Berlin June 2023
' -----------------------------------------------------------------------
Public Trc As clsTrc

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
    ErrSrc = "TestclsTrc." & s
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
    
    Set Trc = New clsTrc
    '~~ Initialization of a new Trace Log File for this Regression test
    With Trc
        .FileName = "RegressionTest_clsTrc.ExecTrace.log"
        .Title = "Regression Test Class Module clsTrc"
        .NewFile
    End With
    
    mErH.Regression = True
        
    mBasic.BoP ErrSrc(PROC), "arg1, arg2"
    Regression_Test_03_Execution_Trace
    Trc.LogInfo = "Test Log-Info explicitly provided"
    Regression_Test_03_Execution_Trace_With_Error

xt: mBasic.EoP ErrSrc(PROC), "arg1, arg2"
    mErH.Regression = False
    Trc.Dsply
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
    Regression_Test_01_1_BoP_missing_TestProc_01a ' ... but this one will be instead
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_01_1_BoP_missing_TestProc_01a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_01_1_BoP_missing_TestProc_01a"
    
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
    Regression_Test_01_2_BoP_missing_TestProc_01a ' ... but this one will be instead
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_01_2_BoP_missing_TestProc_01a()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_01_2_BoP_missing_TestProc_01a"
    
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
    Regression_Test_02_BoP_EoP_TestProc_02a_missing_BoP
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_02a_missing_BoP()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_02a_missing_BoP"
    
'    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_02b_paired_BoP_EoP
    Regression_Test_02_BoP_EoP_TestProc_02d_missing_EoP

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_02b_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_02b_paired_BoP_EoP"
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_02c_paired_BoP_EoP
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_02c_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_02c_paired_BoP_EoP"
    
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

Private Sub Regression_Test_02_BoP_EoP_TestProc_02d_missing_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_02d_missing_EoP"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_02e_BoC_EoC
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_02e_BoC_EoC()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_02e_BoC_EoC"
    
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
    Regression_Test_03_Execution_Trace_TestProc_03a arg1:="xxxx", arg2:="yyyy", arg3:=12.8

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_03a(ByVal arg1 As Variant, _
                                                            ByVal arg2 As Variant, _
                                                            ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_03a"
    
    mBasic.BoP ErrSrc(PROC), arg1 & ", arg2=" & arg2 & ", " & arg3
    Trc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_TestProc_03b
    Regression_Test_03_Execution_Trace_TestProc_03c
    Trc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_03b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_03b"
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

Private Sub Regression_Test_03_Execution_Trace_TestProc_03c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_03c"
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
    Regression_Test_03_Execution_Trace_With_Error_TestProc_03a
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_03a()

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_03a"
    
    mBasic.BoP ErrSrc(PROC)
    Trc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_With_Error_TestProc_03b
    Regression_Test_03_Execution_Trace_With_Error_TestProc_03c
    Trc.EoC ErrSrc(PROC) & " call of 6b and 6c"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_03b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_03b"
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

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_03c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_03c"
    On Error GoTo eh

    mBasic.BoP ErrSrc(PROC)
    '~~ The VB Runtime error 6 is anticipated thus regarded asserted
    '~~ when mErH.Regression = True for this test (set with the
    '~~ calling procedure) the display of the error is suspended
    mErH.Asserted 6 ' VB-Runtime-error overflow
    Trc.LogInfo = "This procedure will raise a VB-Runtime Error 6 (oderflow) - not displayed with the regression test because asserted by mErH.Assert"
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
    
    sFileName = "ExecTrace.log"
    sPath = ThisWorkbook.Path
    sFileFullName = sPath & "\" & sFileName ' creates the file

    '~~ 1. Defaults
    With New clsTrc
        Debug.Assert .FileFullName = sFileFullName
        .Path = sPath
        Debug.Assert .FileFullName = sFileFullName
        .FileName = sFileName
        Debug.Assert .FileFullName = sFileFullName
    End With
    
            
    s = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "ExecTrace.RegressionTest.log")
    If fso.FileExists(s) Then fso.DeleteFile s
    
    '~~ 1. NewFile test
    With New clsTrc
        .NewFile
        Debug.Assert fso.FileExists(.FileFullName)
    End With

    '~~ 2. Go with user-spec log-file (existing default is deleted)
    With New clsTrc
        s = .FileFullName ' the default
        Debug.Assert fso.FileExists(s)
        .FileName = "xxx.log"
        Debug.Assert Not fso.FileExists(.FileFullName)
    End With
    
    Set fso = Nothing
    
End Sub




Attribute VB_Name = "mTestmTrc"
Option Explicit
' -----------------------------------------------------------------------
' Standard Module TestmTrc: Provides all test obligatory being executed
' ========================= when the mTrc code is modified.
'
'
' Uses (for test only): mBasic, fMsg/mMsg, mErH, mTrc
'
' W. Rauschenberger Berlin June 2013
' -----------------------------------------------------------------------
Private FSo         As New FileSystemObject

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

Private Sub BoC(ByVal b_id As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Bnd-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.BoC b_id, b_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.BoC b_id, b_args
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, _
      Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Sub EoC(ByVal e_id As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.EoC e_id, e_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.EoC e_id, e_args
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2014
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    '~~ About
    ErrDesc = err_dscrptn
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    End If
    '~~ Type of error
    If err_no < 0 Then
        ErrType = "Application Error ": ErrNo = AppErr(err_no)
    Else
        ErrType = "VB Runtime Error ":  ErrNo = err_no
        If err_dscrptn Like "*DAO*" _
        Or err_dscrptn Like "*ODBC*" _
        Or err_dscrptn Like "*Oracle*" _
        Then ErrType = "Database Error "
    End If
    
    '~~ Title
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")
    '~~ Description
    ErrText = "Error: " & vbLf & ErrDesc
    '~~ About
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTestmTrc." & s
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
    
#If mTrc = 1 Then
    On Error GoTo eh
    
    mErH.Regression = True
    Set TestAid = Nothing: Set TestAid = New clsTestAid
    TestAid.ModeRegression = True

    Prepare "mTrc"
    '~~ Initialization of a new Trace Log File for this Regression test
    mTrc.NewFile TestAid.TestFolder & "\RegressionTest_mTrc.log"
    mTrc.Title = "Regression Test Module mTrc"
    
    BoP ErrSrc(PROC), "arg1, arg2"
    mTestmTrc.Test_00_DefaultVersusSpecifiedLogFile
    mTrc.FileFullName = TestAid.TestFolder & "\RegressionTest_mTrc.log"
    mTestmTrc.Test_01_Execution_Trace
    mTrc.LogInfo = "Next will raise an error with display bypassed when Regression mode!"
    mTestmTrc.Test_02_Execution_Trace_With_Error
'    mTestmTrc.Test_03_BoP_EoP

xt: EoP ErrSrc(PROC), "arg1, arg2"
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
#Else
    MsgBox "Cond. Comp. Arg" & vbLf & vbLf & _
           "clsTrc = 0: mTrc = 1" & vbLf & vbLf & _
           "is required for testing this component!", vbCritical, "Regression test of Standard Module  ""mTrc"" is not possible!"
#End If
End Sub

Public Sub Test_00_DefaultVersusSpecifiedLogFile()
    Const PROC = "Test_00_DefaultVersusSpecifiedLogFile"
    
    Dim s As String

    s = TestAid.TestFolder & "\TestExecmTrc.log"
    Prepare "mTrc"
    With TestAid
        .TestId = "00-1"
        .Verification = "No file specified rsults in default name"
        .ResultExpected = mTrc.DefaultFileName
        .Result = mTrc.FileFullName
        ' ==============================================================================
        
        .TestId = "00-2"
        .Verification = "NewFile without having specified one - the full name is the default file name"
        .ResultExpected = s
        mTrc.NewFile s
        .Result = mTrc.FileFullName
        ' ==============================================================================
        
        .TestId = "00-3"
        .Verification = "Start with default, change to specified, default is deleted when existing"
        mTrc.NewFile ' becomes the default file
        .ResultExpected = s
        mTrc.FileFullName = s
        .Result = mTrc.FileFullName
        ' ==============================================================================
    
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

End Sub

Public Sub Test_03_BoP_EoP()
' ---------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' ---------------------------------------------------
    Const PROC = "Test_03_BoP_EoP"
    
    Prepare "mTrc"
    BoP ErrSrc(PROC)
    With TestAid
        .TestId = "03-1"
        .Verification = "BoP/EoP missing"
        Test_03_BoP_EoP_TestProc_03a_missing_BoP
        ' ==============================================================================
    
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_BoP_EoP_TestProc_03a_missing_BoP()
' -----------------------------------------------------------
' The error handler is trying its very best not to confuse
' with unpaired BoP/EoP code lines. However, it depends at
' which level this is the case.
' -----------------------------------------------------------
    Const PROC = "Test_03_BoP_EoP_TestProc_03a_missing_BoP"
    
'    BoP ErrSrc(PROC)
    Test_03_BoP_EoP_TestProc_03b_paired_BoP_EoP
    Test_03_BoP_EoP_TestProc_03d_missing_EoP

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_BoP_EoP_TestProc_03b_paired_BoP_EoP()
    Const PROC = "Test_03_BoP_EoP_TestProc_03b_paired_BoP_EoP"
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_03_BoP_EoP_TestProc_03c_paired_BoP_EoP
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_BoP_EoP_TestProc_03c_paired_BoP_EoP()
    Const PROC = "Test_03_BoP_EoP_TestProc_03c_paired_BoP_EoP"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    BoC ErrSrc(PROC) & " trace of some code lines (EoC statement missing!)"

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_BoP_EoP_TestProc_03d_missing_EoP()
    Const PROC = "Test_03_BoP_EoP_TestProc_03d_missing_EoP"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Test_03_BoP_EoP_TestProc_03e_BoC_EoC
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_BoP_EoP_TestProc_03e_BoC_EoC()
    Const PROC = "Test_03_BoP_EoP_TestProc_03e_BoC_EoC"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
        
    Dim i As Long: Dim j As Long: j = 10000000
    BoC PROC & " code trace empty loop 1 to " & j
    For i = 1 To j
    Next i
    EoC PROC & " code trace empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_01_Execution_Trace"
    On Error GoTo eh
    
    Prepare "mTrc"
    BoP ErrSrc(PROC)
    With TestAid
        .TestId = "01-1"
        .Verification = "Full trace test"
        Test_01_Execution_Trace_TestProc_01a arg1:="xxxx", arg2:="yyyy", arg3:=12.8
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_Execution_Trace_TestProc_01a(ByVal arg1 As Variant, _
                                                 ByVal arg2 As Variant, _
                                                 ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Test_01_Execution_Trace_TestProc_01a"
    
    BoP ErrSrc(PROC), arg1 & ", arg2=" & arg2 & ", " & arg3
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_01_Execution_Trace_TestProc_01b
    Test_01_Execution_Trace_TestProc_01c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_Execution_Trace_TestProc_01b()
    
    Const PROC = "Test_01_Execution_Trace_TestProc_01b"
    On Error GoTo eh

    BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_Execution_Trace_TestProc_01c()
    
    Const PROC = "Test_01_Execution_Trace_TestProc_01c"
    On Error GoTo eh

    BoP ErrSrc(PROC)

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_Execution_Trace_With_Error()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_02_Execution_Trace_With_Error"
    On Error GoTo eh
    
    Prepare "mTrc"
    BoP ErrSrc(PROC)
    With TestAid
        .TestId = "02-1"
        .Verification = "Execution trace with error"
        Test_02_Execution_Trace_With_Error_TestProc_01a
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_02_Execution_Trace_With_Error_TestProc_01a()

    On Error GoTo eh
    Const PROC = "Test_02_Execution_Trace_With_Error_TestProc_01a"
    
    BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_02_Execution_Trace_With_Error_TestProc_01b
    Test_02_Execution_Trace_With_Error_TestProc_01c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_02_Execution_Trace_With_Error_TestProc_01b()
    
    Const PROC = "Test_02_Execution_Trace_With_Error_TestProc_01b"
    On Error GoTo eh

    BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_02_Execution_Trace_With_Error_TestProc_01c()
    
    Const PROC = "Test_02_Execution_Trace_With_Error_TestProc_01c"
    On Error GoTo eh

    BoP ErrSrc(PROC)
    '~~ The VB Runtime error 6 is anticipated thus regarded asserted
    '~~ when mErH.Regression = True for this test (set with the
    '~~ calling procedure) the display of the error is suspended
    mErH.Asserted 6 ' VB-Runtime-error overflow
    mTrc.LogInfo = "This procedure will raise a VB-Runtime Error 6 (oderflow) - not displayed with the regression test because asserted by mErH.Assert"
    Dim i As Long
    i = i / 0 ' Error !!!!

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub




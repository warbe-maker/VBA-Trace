Attribute VB_Name = "TestmTrc"
Option Explicit
' -----------------------------------------------------------------------
' Standard Module TestmTrc: Provides all test obligatory being executed
' ========================= when the mTrc code is modified.
'
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

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.        | Installed component |
' |-------------------------|---------------------|
' | XcTrc_mTrc = 1          | mTrc                |
' | XcTrc_clsTrc = 1        | clsTrc              |
' | ErHComp = 1             | mErH                |
' I.e. both components are independant from each other!
' Note: This procedure is obligatory for any VB-Component using either the
'       the 'Common VBA Error Services' and/or the 'Common VBA Execution Trace
'       Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ";")

#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf XcTrc_clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf XcTrc_mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         | ErHComp = 1             | mErH                |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf XcTrc_clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf XcTrc_mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option
' (Cond. Comp. Arg. 'Debugging = 1') and an optional additional
' "about the error" information which may be connected to an error message by
' two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Cond. Comp. Arg. 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Cond. Comp. Arg. 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
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
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
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
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "TestmTrc." & s
End Function

Public Sub Regression_Test()
' ----------------------------------------------------------------------------
' Attention!
' ==========
' Testing the mTrc module requires the Cond. Comp. Arg.:
' - XcTrc_mTrc = 1
' - XcTrc_clsTrc = 0
'
' The execution trace writes a trace-log file which defaults to to the
' ThisWorkbook's name with a .trc extentsion and is located a the
' ThisWorkbook's parent folder. With this regression test the .trc file's
' content is displayed by the mTrc.Dsply service by means of ShellRun.
' ------------------------------------------------------------------------------
    Const PROC = "Regression_Test"
    
    On Error GoTo eh
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    mTrc.NewLog
    mTrc.FileFullName = ThisWorkbook.Path & "\RegressionTest.log"
    mTrc.LogTitle = "Regression Test module mTrc"
        
    mErH.Regression = True
        
    BoP ErrSrc(PROC)
    Regression_Test_03_Execution_Trace
    mTrc.LogInfo = "Test Log-Info explicitly provided"
    Regression_Test_03_Execution_Trace_With_Error

xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
'    BoP ErrSrc(PROC) this procedure will not be recognized as "Entry Procedure" ...
    Regression_Test_01_1_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    Regression_Test_01_2_BoP_missing_TestProc_1a ' ... but this one will be instead
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
'    BoP ErrSrc(PROC)
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1a_missing_BoP
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
'    BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP
    Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1b_paired_BoP_EoP"
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1c_paired_BoP_EoP"
    
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

Private Sub Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1d_missing_EoP"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC()
    Const PROC = "Regression_Test_02_BoP_EoP_TestProc_1e_BoC_EoC"
    
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

Public Sub Regression_Test_03_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Regression_Test_03_Execution_Trace"
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Regression_Test_03_Execution_Trace_TestProc_6a arg1:="xxxx", arg2:="yyyy", arg3:=12.8

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_6a(ByVal arg1 As Variant, _
                                               ByVal arg2 As Variant, _
                                               ByVal arg3 As Variant)

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6a"
    
    BoP ErrSrc(PROC), arg1, "arg2=", arg2, arg3
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_TestProc_6b
    Regression_Test_03_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_TestProc_6b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6b"
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

Private Sub Regression_Test_03_Execution_Trace_TestProc_6c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    BoP ErrSrc(PROC)

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    BoP ErrSrc(PROC)
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6a
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6a"
    
    BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6b
    Regression_Test_03_Execution_Trace_With_Error_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6b()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6b"
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

Private Sub Regression_Test_03_Execution_Trace_With_Error_TestProc_6c()
    
    Const PROC = "Regression_Test_03_Execution_Trace_With_Error_TestProc_6c"
    On Error GoTo eh

    '~~ The VB Runtime error 6 is anticipated thus regarded asserted
    '~~ when mErH.Regression = True for this test (set with the
    '~~ calling procedure) the display of the error is suspended
    mErH.Asserted 6
    mTrc.LogInfo = "This is just an info"
    Dim i As Long
    i = i / 0 ' Error !!!!

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    
    '~~ 2. Clear test
    mTrc.NewLog
    Debug.Assert fso.FileExists(mTrc.FileFullName)
    
    '~~ 3. Go with user-spec log-file (existing default is deleted)
    mTrc.Initialize ' setup defaults
    mTrc.FileName = "ExecTrace.My.log"
    Debug.Assert mTrc.FileFullName = sPath & "\ExecTrace.My.log"
    Debug.Assert Not fso.FileExists(mTrc.FileFullName)
    
    '~~ 4. Cleanup
    mTrc.NewLog
    Debug.Assert fso.FileExists(mTrc.FileFullName)
    
    Set fso = Nothing
    
End Sub


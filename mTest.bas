Attribute VB_Name = "mTest"
Option Explicit
' -----------------------------------------------------------------------
' Standard Module mTest: Provides all test procedures obligatiory after
'                        any code modification.
' Note: The tested mTrc module must not use the error handler module mErH
'       because it itself is an optional module for this erro handler.
'       In replacing this error handler all resources for a local error
'       handling had been copied from the mErH module.
'
' W. Rauschenberger, Berlin Nov 2020
' -----------------------------------------------------------------------
Private Const TYPE_APP_ERR  As String = "Application error "
Private Const TYPE_VB_ERR   As String = "VB Runtime error "
Private Const TYPE_DB_ERROR As String = "Database error "
Private Const CONCAT = "||"
' ----------------------------------------------------------------------
' Deklarations for the use of the fMsg UserForm (Alternative VBA MsgBox)
' Note: These declarations are part of the mErH module when the mTrc
'       module is used as the optional error handler module.
' ----------------------------------------------------------------------
Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type                            ' -------------------
' end of fMsg declarations ---------------------------------------------
Private Enum enErrorType
    VBRuntimeError
    ApplicationError
    DatabaseError
End Enum

Private bRegressionTest As Boolean

Public Function AppErr(ByVal errno As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application
' Error Number and vbObjectErrror is added to it into a negative
' number in order not to confuse with a VB runtime error.
' When the error number <l> is negative it is considered an
' Application Error and vbObjectError is added to convert it back
' into its origin positive number.
' ------------------------------------------------------------------
    If errno < 0 Then
        AppErr = errno - vbObjectError
    Else
        AppErr = vbObjectError + errno
    End If
End Function

Private Sub ErrMsg(ByVal errno As Long, _
                   ByVal errsource As String, _
                   ByVal errdscrptn As String, _
                   ByVal errline As Long)
' ----------------------------------------------
'
' ----------------------------------------------
    Dim sErrInfo As String
    
    MsgBox Prompt:="Error description" & vbLf & _
                    err.Description, _
           buttons:=vbOKOnly, _
           Title:="VB Runtime error " & errno & " in " & errsource & IIf(errline <> 0, " at line " & errline, "")
#If ExecTrace Then
    '~~ Any other error handling but the Common VBA Error Handler (module mErH) will finish the execution trace
    '~~ in case of an error explicitely which will display the trace result if any
    mTrc.Finish ErrorDetails(errnumber:=errno, errsource:=errsource, sErrLine:=ErrorLine(errline:=errline))
    mTrc.Terminate ' clean up
#End If
End Sub

Private Function ErrorDetails( _
                 ByVal errnumber As Long, _
                 ByVal errsource As String, _
                 ByVal sErrLine As String) As String
' --------------------------------------------------
' Returns the kind of error, the error number, and
' the error line (if available) as string.
' --------------------------------------------------
    
    Select Case ErrorType(errnumber, errsource)
        Case ApplicationError:              ErrorDetails = ErrorTypeString(ErrorType(errnumber, errsource)) & AppErr(errnumber)
        Case DatabaseError, VBRuntimeError: ErrorDetails = ErrorTypeString(ErrorType(errnumber, errsource)) & errnumber
    End Select
        
    If sErrLine <> vbNullString Then ErrorDetails = ErrorDetails & ErrorLine(Erl)

End Function

Private Function ErrorLine( _
                 ByVal errline As Long) As String
' -----------------------------------------------
' Returns a complete errol line message.
' -----------------------------------------------
    If errline <> 0 _
    Then ErrorLine = " (at line " & errline & ")" _
    Else ErrorLine = vbNullString
End Function

Private Function ErrorType( _
                 ByVal errnumber As Long, _
                 ByVal errsource As String) As enErrorType
' --------------------------------------------------------
' Return the kind of error considering the error source
' (errsource) and the error number (errnumber).
' --------------------------------------------------------

   If InStr(1, errsource, "DAO") <> 0 _
   Or InStr(1, errsource, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, errsource, "ODBC") <> 0 _
   Or InStr(1, errsource, "Oracle") <> 0 Then
      ErrorType = DatabaseError
   Else
      If errnumber > 0 _
      Then ErrorType = VBRuntimeError _
      Else ErrorType = ApplicationError
   End If
   
End Function

Private Function ErrorTypeString(ByVal errtype As enErrorType) As String
    Select Case errtype
        Case ApplicationError:  ErrorTypeString = TYPE_APP_ERR
        Case DatabaseError:     ErrorTypeString = TYPE_DB_ERROR
        Case VBRuntimeError:    ErrorTypeString = TYPE_VB_ERR
    End Select
End Function

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
    Test_1_Execution_Trace
    Test_2_Execution_Trace_With_Error

xt: mTrc.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Public Sub Test_1_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_1_Execution_Trace"
    On Error GoTo eh
'    mTrc.DisplayedInfo = Compact
    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " Code trace"
    Test_1_Execution_Trace_TestProc_6a
    mTrc.EoC ErrSrc(PROC) & " Code trace"
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Test_1_Execution_Trace_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Test_1_Execution_Trace_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_1_Execution_Trace_TestProc_6b
    Test_1_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Sub Test_1_Execution_Trace_TestProc_6b()
    
    Const PROC = "Test_1_Execution_Trace_TestProc_6b"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 100000
        s = Application.Path ' to produce some execution time
    Next i
    
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Test_1_Execution_Trace_TestProc_6c()
    
    Const PROC = "Test_1_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    ' produce some execution time
    Dim i As Long: Dim s As String
    For i = 1 To 10000
        s = Application.Name
    Next i

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Public Sub Test_2_Execution_Trace_With_Error()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_2_Execution_Trace_With_Error"
    On Error GoTo eh
'    mTrc.DisplayedInfo = Compact
    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " Code trace"
    Test_2_Execution_Trace_With_Error_TestProc_6a
    mTrc.EoC ErrSrc(PROC) & " Code trace"
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Sub Test_2_Execution_Trace_With_Error_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Test_2_Execution_Trace_With_Error_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_2_Execution_Trace_With_Error_TestProc_6b
    Test_2_Execution_Trace_With_Error_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Sub Test_2_Execution_Trace_With_Error_TestProc_6b()
    
    Const PROC = "Test_2_Execution_Trace_With_Error_TestProc_6b"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    Dim i As Long
    Dim s As String
    For i = 1 To 10000
        s = Application.Path ' to produce some execution time
    Next i
    
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub

Private Sub Test_2_Execution_Trace_With_Error_TestProc_6c()
    
    Const PROC = "Test_2_Execution_Trace_With_Error_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)
    
    ' produce some execution time
    Dim i As Long: Dim s As String
    For i = 1 To 10000
        s = Application.Name
    Next i
    i = i / 0
    
xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
End Sub


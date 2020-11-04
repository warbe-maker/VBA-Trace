Attribute VB_Name = "mTest"
Option Explicit

Private Const CONCAT = "||"
' ----------------------------------------------------------------------
' Deklarations for the use of the fMsg UserForm (Alternative VBA MsgBox)
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
    Test_6_Execution_Trace

xt: mTrc.EoP ErrSrc(PROC)
    bRegressionTest = False
    Exit Sub
    
eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Public Sub Test_6_Execution_Trace()
' ------------------------------------------------------
' White-box- and regression-test procedure obligatory
' to be performed after any code modification.
' Display of an execution trace along with this test
' requires a conditional compile argument ExecTrace = 1.
' ------------------------------------------------------
    
    Const PROC = "Test_6_Execution_Trace"
    On Error GoTo eh
'    mTrc.DisplayedInfo = Compact
    mTrc.DisplayedInfo = Detailed
    
    mTrc.BoP ErrSrc(PROC)
    Test_6_Execution_Trace_TestProc_6a
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6a()

    On Error GoTo eh
    Const PROC = "Test_6_Execution_Trace_TestProc_6a"
    
    mTrc.BoP ErrSrc(PROC)
    mTrc.BoC ErrSrc(PROC) & " call of 6b and 6c"
    Test_6_Execution_Trace_TestProc_6b
    Test_6_Execution_Trace_TestProc_6c
    mTrc.EoC ErrSrc(PROC) & " call of 6b and 6c"
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6b()
    
    Const PROC = "Test_6_Execution_Trace_TestProc_6b"
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
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Test_6_Execution_Trace_TestProc_6c()
    
    Const PROC = "Test_6_Execution_Trace_TestProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub ErrMsg(ByVal errno As Long, _
                   ByVal errsource As String, _
                   ByVal errdscrptn As String, _
                   ByVal errline As Long)
' ----------------------------------------------
'
' ----------------------------------------------
    MsgBox Prompt:="Error description" & vbLf & _
                    err.Description, _
           buttons:=vbOKOnly, _
           Title:="VB Runtime error " & errno & " in " & errsource & IIf(errline <> 0, " at line " & errline, "")
End Sub


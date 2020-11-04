Attribute VB_Name = "mDemo"
Option Explicit

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
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6a()

    Const PROC = "Demo_6_Execution_Trace_DemoProc_6a"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    Demo_6_Execution_Trace_DemoProc_6b
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6b()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6b"
    On Error GoTo eh
    
    mTrc.BoP ErrSrc(PROC)
    
    Demo_6_Execution_Trace_DemoProc_6c
    
    Dim i As Long: Dim j As Long: j = 10000000
    mTrc.BoC PROC & " empty loop 1 to " & j
    For i = 1 To j
    Next i
    mTrc.EoC PROC & " empty loop 1 to " & j ' !!! the string must match with the BoC statement !!!
    
    mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Sub Demo_6_Execution_Trace_DemoProc_6c()
    
    Const PROC = "Demo_6_Execution_Trace_DemoProc_6c"
    On Error GoTo eh

    mTrc.BoP ErrSrc(PROC)

xt: mTrc.EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg err.Number, ErrSrc(PROC), err.Description, Erl
#If Debugging Then
    Stop: Resume
#End If
End Sub

Private Function ErrSrc(ByVal s As String) As String
' ---------------------------------------------------
' Prefix procedure name (s) by this module's name.
' Attention: The characters > and < must not be used!
' ---------------------------------------------------
    ErrSrc = "mDemo." & s
End Function

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

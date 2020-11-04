Attribute VB_Name = "mAnyModule"
Option Explicit
' -------------------------------------------------------------------------------
' Sample module when using mErrHndlr
' -------------------------------------------------------------------------------
Const MODNAME = "mAnyModule" ' Module name for error handling and execution trace

Private Sub AnyProc()
' -------------------------------------------------------------------------------
' Sample procedure using mErrHndlr
' -------------------------------------------------------------------------------
Const PROC As String = "AnyProc" ' This procedure's name for error handling and execution trace

    On Error GoTo eh
    mTrc.BoP ErrSrc(PROC) ' Begin of Procedure (push stack and begin of execution trace)

    ' any code

xt: ' any "finally" code
    mTrc.EoP ErrSrc(PROC) ' End of Procedure (pop stack and end of execution trace)
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Split(ThisWorkbook.Name, ".")(0) & "." & MODNAME & "." & sProc
End Function

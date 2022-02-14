# Common VBA Execution Trace

## Service
Logs execution trace entries in a file which defaults to _ExecTrace.log_ in _ThisWorkbook's_ parent folder. The log file content:
![](assets/ExecutionTrace.png)

For details about the individual services of the component see the inline documentation.

## Installation
1. Download [mTrc.frm][4] and import it to your VB-Project
2. **Optionally** download [mMsg.bas][3], [fMsg.frm][1], and [fMsg.frx][2] and import mMsg.bas and fMsg.frm. To activate these components usage set the _Conditional Compile Argument_ `MsgComp = 1 ` This will enable the mTrc.Dsply service to display the trace log result. When these means are not provided with the VB-Project the trace log file will need to be displayed by any tool of the users choice.

## Usage
Copy the following code into any module in which there will be a ['to-be-traced' procedure](#to-be-traced-procedure) to ensure a unique identification of any procedure by prefixxing it with the module's name:
```vbs
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "<module-name>." & sProc
End Function
```
and adjust the <module-name>!

The following procedures not only will keep the use of the _mTrc_ component **optional** but also the _mErH_ and the _mMsg_ component.

```vbs
Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' indication.
' - Serves for a comprehensive display of an error message when the Common VBA
'   Error Services Component mErH is installed and the Conditional Compile
'   Argument 'ErHComp = 1'
' - Serves for the execution trace when the Common VBA Execution Trace Service
'   Component mTrc is installed and the Conditional Compile Argument
'   'ExecTrace = 1'.
' - May alternatively be copied as a Private procedure into any component when
'   this mBasic component is not installed but the mErH and or mTrc services
'   are
' Note: Because the error handling also hands over an 'End of Procedure' to the
'       mTrc component (when installed and 'ExecTrace = 1') an explicit call of
'       mTrc.EoP is only performed when mErH is not installed/in use.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' indication.
' - Serves for a comprehensive display of an error message when the Common VBA
'   Error Services Component mErH is installed and the Conditional Compile
'   Argument 'ErHComp = 1'
' - Serves for the execution trace when the Common VBA Execution Trace Service
'   Component mTrc is installed and the Conditional Compile Argument
'   'ExecTrace = 1'.
' - May alternatively be copied as a Private procedure into any component when
'   this mBasic component is not installed but the mErH and or mTrc services
'   are
' Note: Because the error handling also hands over an 'End of Procedure' to the
'       mTrc component (when installed and 'ExecTrace = 1') an explicit call of
'       mTrc.EoP is only performed when mErH is not installed/in use.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

```

### To be traced procedure
The following code lines will trace a procedures execution provided the _Conditional Compile Argument_ `ExecTrace = 1`:
```vbs
Private Sub Any()
    Const PROC = "Any"
    '...
    
    BoP ErrSrc(PROC)
    ' any code lines

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: ' any error handling
End Sub
```

> ***Hint 1:*** Avoid using **`Exit ...`** to terminate a procedure's execution but use ***`Goto xt`*** instead to ensure the EoP (end of procedure) statement is not bypassed.<br>
***Hint 2:*** An error handling should preferably end with a ***`Goto xt`*** in order to provide a 'clean exit'.


### Execution Trace (mTrc) and Error Services (mErH)
This _Common VBA Execution Trace Component (mTrc)_ and the _Common VBA Error Services Component (mErH)_ have the following in common:
1. Both use in each component/module the `ErrSrc` function to uniquely identify a procedure's name (i.e. prefix it with the component's name)
3. Both use BoP/EoP statements to indicate the <u>B</u>egin and <u>E</u>nd <u>o</u>f a <u>P</u>rocedure.<br>The execution trace uses the statements to begin/end the trace of a procedure<br>the error uses the statements to indicate an 'entry procedure' to which the error is passed on for being displayed (which allows gathering the 'path to the error'.

### Me and the public
I do not like the idea to maintain different versions of a component, one for being used in my own VB-Projects an another 'public' version. To achieve this at first I try to keep components as autonomous as possible and second I keep those components which are obligatory for me personal, optional for others. This is achieved by a couple of procedures I add to any component and by the use of the _Conditional Compile Arguments_ 'Debugging, ExecTrace, MsgComp, and ErHComp. By these means other users are bothered by my personal preferences as little as possible.

### Contribution
Contribution of any kind in any form is welcome - preferably by raising an issue.


[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[5]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[6]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas

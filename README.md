## Common VBA Execution Trace Services

Writes records of traced executions of procedures and code snippets to a file which defaults to _ExecTrace.log_ in _ThisWorkbook's_ parent folder. Example of a log file's content:
![](assets/ExecutionTrace.png)

For details about the individual services of the component see the inline documentation.

## Services
| Service    | Purpose |
| ---------- | ------- |
|            |         |
| _BoC_      | Indicates the (B)egin (o)f the execution trace of a (C)ode snippet. |
| _BoP_      | Indicates the (B)egin (o)f the execution trace of a (P)rocedure. |
| _BoP\_ErH_ | Exclusively used by the mErH module.
| _Continue_ | Commands the execution trace to continue taking the execution time when it had been paused. Pause and Continue is used by the mErH module for example to avoid useless execution time taking while waiting for the users reply.|
| _Dsply_     | Displays the content of the trace log file. Available only when the mMsg/fMsg modules are installed and this is indicated by the Conditional Compile Argument 'MsgComp = 1'. Without mMsg/fMsg the trace result log will be viewed with any appropriate text file viewer. |
| _EoC_       | Indicates the (E)nd (o)f the execution trace of a (C)ode snippet. |
| _EoP_       | Indicates the (E)nd (o)f the execution trace of a (P)rocedure. |
| _Pause_     | Stops the execution traces time taking, e.g. while an error message is displayed. |
| _LogFile _  | Get/Let property for the full name of a desired trace log file which defaults to "ExecTrace.log" in ThisWorkbook's parent folder.
| _LogInfo_   | Adds an entry to the trace log file by considering the current nesting level (i.e. the indentation). |

## Installation
Download [mTrc.frm][4] and import it to your VB-Project.

## Usage
Copy the following code into any ['to-be-traced' procedure](#to-be-traced-procedure).

| Procedure | Purpose |
| --------- | ------- |
| _ErrSrc_  | Ensures a a unique identification of any procedure by prefixing it with the adjusted! module's name (will also be used for the [Common VBA Error Services][7] when installed) |
| _BoP\/EoP_ | Keeps the availability of the _mTrc_ module optional. Will also serve for the (will also be used for the [Common VBA Error Services][7] when installed) when installed. |

```vb
Private Function ErrSrc(ByVal proc_name As String) As String
    ErrSrc = "<module-name>." & proc_name
End Function

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
    On Error Goto eh
    '...
    
    BoP ErrSrc(PROC)
    ' any code lines

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: .... ' any error handling
    Goto xt ' clean exit
End Sub
```

> Avoid using **`Exit ...`** to terminate a procedure's execution but use ***`Goto xt`*** instead to ensure the EoP (end of procedure) statement is not bypassed.<br>

> An error handling should preferably end with a ***`Goto xt`*** in order to provide a 'clean exit'.

### Personal and public use of (my) _Common Components_
I do not like the idea maintaining different code versions of _Common Components_, one which I use in my VB-Projects and another 'public' version. On the other hand I do not want to urge users of my _Common Components_ to also use the other _Common Components_ which have become a de facto standard for me.

#### Managing the splits
My primary goal is to provide _Common Components_ which function as autonomous as possible - and also to optionally use them together with the/my [Common VBA Message Services][5] and the [Common VBA Error Services][7]. This 'optionally installed' is primarily achieved by the use of a couple of _Conditional Compile Arguments_ and procedures also by a couple of procedures which only optionally use other _Common Components_ only when installed.

| Conditional Compile Argument | Purpose |
| ---------------------------- | ------- |
| _Debugging_                  | Indicates that error messages should be displayed with a debugging option allowing to resume the error line |
| _ExecTrace_                  | Indicates that the _[mTrc][4]_ module is installed
| _MsgComp_                    | indicates that the _[mMsg][3]_, _[fMsg.frm][1]_, and _[fMsg.frx][2]_ are installed |
| _ErHComp_                    | Indicates that the _[mErH][6]_ is installed |

By these means other users are no bothered by my personal preferences - or are only as little as possible :-).

#### Execution Trace _(mTrc)_ and Error Services _(mErH)_
This _Common VBA Execution Trace Component (mTrc)_ and the _Common VBA Error Services Component (mErH)_ have the following in common:
1. Both use in each component/module the `ErrSrc` function to uniquely identify a procedure's name (i.e. prefix it with the component's name)
3. Both use _BoP/EoP_ statements to indicate the <u>B</u>egin and <u>E</u>nd <u>o</u>f a <u>P</u>rocedure.<br>The execution trace uses the statements to begin/end the trace of a procedure<br>the error uses the statements to indicate an 'entry procedure' to which the error is passed on for being displayed (which allows gathering the 'path to the error'.
### Contribution
Contribution of any kind in any form is welcome - preferably by raising an issue.


[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[5]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[6]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[7]:https://github.com/warbe-maker/Common-VBA-Error-Services
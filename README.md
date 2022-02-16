## Common VBA Execution Trace Services

Writes records of traced executions of procedures and code snippets to a file which defaults to _ExecTrace.log_ in _ThisWorkbook's_ parent folder. Example of a log file's content:
![](assets/ExecutionTrace.png)

For details about the individual services of the component see the inline documentation.

## Services
| Service    | Purpose |
| ---------- | ------- |
|            |         |
| _BoC_      | Indicates the (B)egin (o)f the execution trace of a (C)ode snippet. Does nothing unless the Conditional Compile Argument `ExecTrace = 1` |
| _BoP_      | Indicates the (B)egin (o)f the execution trace of a (P)rocedure. Does nothing unless the Conditional Compile Argument `ExecTrace = 1` |
| _BoP\_ErH_ | Exclusively used by the mErH module.
| _Continue_ | Commands the execution trace to continue taking the execution time when it had been paused. Pause and Continue is used by the mErH module for example to avoid useless execution time taking while waiting for the users reply.|
| _Dsply_     | Displays the content of the trace log file. Available only when the mMsg/fMsg modules are installed and this is indicated by the Conditional Compile Argument 'MsgComp = 1'. Without mMsg/fMsg the trace result log will be viewed with any appropriate text file viewer. |
| _EoC_       | Indicates the (E)nd (o)f the execution trace of a (C)ode snippet. Does nothing unless the Conditional Compile Argument `ExecTrace = 1` |
| _EoP_       | Indicates the (E)nd (o)f the execution trace of a (P)rocedure. Does nothing unless the Conditional Compile Argument `ExecTrace = 1` |
| _Pause_     | Stops the execution traces time taking, e.g. while an error message is displayed. |
| _LogFile_  | Property, string expression<br>- Let: Specifies the full name of a desired trace log file, defaults to "ExecTrace.log" in ThisWorkbook's parent folder when none is specified<br>- Get: Returns the used log file's full name |
| _LogInfo_   | Adds an entry to the trace log file by considering the current nesting level (i.e. the indentation). |
| _LogTitle_  | Property, string expression, Let only, specifies the begin and end trace title, defaults to "Begin execution trace", "End execution trace" |

## Installation
Download [mTrc.frm][4] and import it to your VB-Project.

## Usage
> This _Common Component_ is prepared to function completely autonomously ( download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][8] for more details.

### Module/component preparation
Copy the following code into any module/component a procedure may possibly be traced:

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
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

```

### Procedure (Sub, Function, Property) preparation
The following code lines will trace a procedures execution provided the _Conditional Compile Argument_ `ExecTrace = 1`:
```vbs
Private <kind of procedure> Any()
    Const PROC = "Any"
    On Error Goto eh
    '...
    
    BoP ErrSrc(PROC)
    ' any code lines

xt: EoP ErrSrc(PROC)
    Exit <kind of procedure>
    
eh: .... ' any error handling
    Goto xt ' clean exit
End <kind of procedure>
```

> Avoid using **`Exit <kind of procedure>`** to terminate a procedure's execution but use ***`Goto xt`*** instead to ensure the EoP (<u>E</u>nd <u>o</u>f <u>P</u>rocedure) statement is not bypassed.<br>

> An error handling should preferably end with a ***`Goto xt`*** in order to provide a 'clean exit'.

### Contribution
Contribution of any kind in any form is welcome - preferably by raising an issue.


[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[5]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[6]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[7]:https://github.com/warbe-maker/Common-VBA-Error-Services
[8]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
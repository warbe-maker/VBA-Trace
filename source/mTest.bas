Attribute VB_Name = "mTest"
Option Explicit

Public Trc      As clsTrc
Public TestAid  As New clsTestAid

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
' W. Rauschenberger Berlin, Jan 2024
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
    ErrSrc = "mTest." & s
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add key:=vKey, item:=k_dct.item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Prepare(ByVal p_comp As String)
    
    Dim s As String

    Select Case p_comp
        Case "clsTrc"
            #If clsTrc = 0 Or mTrc = 1 Then
                MsgBox "Cond. Comp. Arg. `clsTrc = 1` And `mTrc = 0` is required for testing the component clsTrc!", vbCritical
            #End If
        Case "mTrc"
            #If clsTrc = 1 Or mTrc = 0 Then
                MsgBox "Cond. Comp. Arg. `clsTrc = 0` And `mTrc = 1` is required for testing the component mTrc!", vbCritical
            #End If
            Set Trc = Nothing
        Case Else
            MsgBox "No 'TestedComp' specified with 'TestAid'!", vbCritical
    End Select
    
    TestAid.TestedComp = p_comp
    
    If Not TestAid.ModeRegression Then
        #If clsTrc = 1 Then
            If Trc Is Nothing Then
                Set Trc = New clsTrc
                s = TestAid.TestFolder & "\TestExecTrace.log"
                Trc.NewFile s
            End If
        #Else
            Set Trc = Nothing
        #End If
        #If mTrc = 1 Then
                Set Trc = Nothing
                mTrc.NewFile s
        #End If
    End If
    
End Sub

Public Sub CondCompArgTest()
    
    Dim cmb     As CommandBar
    Dim cbb     As CommandBarButton
    Dim dct     As New Dictionary
    Dim v       As Variant
    Dim xlApp   As Application
    Dim cbp     As CommandBarPopup
    Dim v2      As Variant
    
    Set xlApp = CreateObject("Excel.Application")

'    For Each cmb In xlApp.VBE.CommandBars
'        If cmb.Index = 1 Then
'            For Each v In cmb.Controls
'                If TypeName(v) = "CommandBarPopup" Then
'                    Set cbp = v
'                    If cbp.id = "30007" Then
'                        For Each v2 In cbp.Controls
'                            Set cbb = v2
'                            dct.Add cmb.Index & "." & cbp.id & "." & cbb.id & "." & cmb.Name & "." & cbp.Caption & "." & cbb.Caption, cbp
'                        Next v2
'                    End If
'                End If
'            Next v
'        End If
'    Next cmb
'
'    KeySort dct
'    For Each v In dct
'        Set cbp = dct(v)
'        Debug.Print v & " : Caption: " & cbp.Caption
'    Next v
    
    TestAid.CondCompArgSet "mTrc = 0"
    
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestAid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' Class Module clsTestAid: Common, tests supporting services, including
' ======================== Regression-Testing whereby each individula test is
' identified by test-number which may be any distinct string.
'
' Regression      In regression test mode application errors asserted
' test mode:      beforehand do not interrupt the test stream. Additionally
'                 test results are logged in a "Regression.Summary.log" file
'                 and may be displayed finally. In "non-regression-test-mode",
'                 when the result and a specified result expected differ,
'                 WinMerge is used to display the difference (an error is
'                 displayed when not installed), provided both, the Result
'                 and the ResultExpected are test files (.
' Test result     The result of a test may be an Array, a Collection, a
' evaluation:     Dictionary, a File, or a String with delimited Sub-Strings.
'                 The "expected result" may be either of the above, whereas a
'                 file will be preferred because it may be prepared beforehand.
'                 When both. the Result and the ResultExpected had been
'                 provided the evaluation is triggered. Both are transformed
'                 into a Collection and compared item by item.
' Variant result  Result  The module uses VarTrans (a copy from mVarTrans) to
' transformation: provide transformation services such as transforming results
'                 into a file in order to make results comparable with
'                 expected results.
'
' Public services:
' ----------------
' DsplySummary    M     Displays a RegressionTest.log file provided
'                       ModeRegression = True.
' ModeRegression  P   w Boolean expression, indicates regression mode. When
'                       True, a failed test result is not displayed but
'                       displayed as summary by the DsplySummary service.
' Result          P r/w Variant expression registers/returns the result of a
'                       test, may be a file object, any other object or any
'                       other data type.
' ResultExpected  P r/w The expected result of a test, a file object, any
'                       other object or any other data type
' TempFile              Saves the name of a temporary test file in a
'                       Collection for being removed when the class terminates
'                       or explicitely by the TestFilesRemove service.
' TestDscrpt      P   w A string describing the purpose of a test.
' TestedProc      P   w The name of the tested procedure (a Sub, Function,
'                       Property, or Method).
' TestFile        P r/w Returns a standard test file name based on a provided
'                       test number or registers a test file provided either
'                       as full name or file object for being removed when
'                       when the class terminates or explicitely by the
'                       TestFilesRemove service.
' TestFilesRemove M     Removes all files provided either by the TestFile
'                       or the TempFile service.
' TestFolder      P r/w String expression, specifies the test folder
' TestNumber      P   w Essential, obligatory distinct identifcation of a
'                       test, provided as the very first Property for each
'                       test.
' TestResult      P r   Returns Passed or Failed
' BoC/EoC         M     Effective only when the execution trace module (mTrc or
'                       clsTrc) is installed, indicates the begin of a code
'                       sequence.
' BoP/EoP         M     Effective only when the execution trace module (mTrc or
'                       clsTrc) is installed, indicates the begin/end of a
'                       procedure.
'
' Requires:
' - Reference to "Microsoft Scripting Runtime"
'
' Uses:
' - clsLog              To write and display failed test details
' - mTrc/clsTrc         Only when an execution trace is desired, indicated by
'                       Cond. Comp. Arg. `mTrc = 1` or `clsTrc = 1`
'
' W. Rauschenberger, Berlin Apr 2024
' ----------------------------------------------------------------------------
Public FSo               As New FileSystemObject

Private Const RESULT_PASSED As String = " Passed "
Private Const RESULT_FAILED As String = " F a i l e d ! "
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias _
"GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SETTEXT = &HC
Private Const BM_CLICK = &HF5

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

Private bModeRegression     As Boolean
Private cllTestFiles        As New Collection
Private dctResult           As New Dictionary
Private dctResultExpected   As New Dictionary
Private dctTestSummary      As New Dictionary
Private lMaxLenTestProc     As Long
Private sRegrTestSummaryLog As String
Private sTestDscrpt         As String
Private sTestedComp         As String
Private sTestedProc         As String
Private sTestedType         As String
Private sTestFolder         As String
Private sTestNumber         As String
Private sResultEvaluated    As String
Private sSubTestId          As String
Private Log                 As clsLog
Private v                   As Variant
Private sDelim              As String
Private sTestFile           As String

Private Sub Class_Initialize()
    
    '~~ Provide the default test folder
    TestFolder = ThisWorkbook.Path & "\Test"
    
    '~~ Provide a new default regression test summary log file's full name
    RegrTestSummaryLog = sTestFolder & "\Regression.Summary.log"
    
    Set cllTestFiles = New Collection
    sSubTestId = 1
    Set Log = New clsLog
    
End Sub

Private Sub Class_Terminate()

    Set dctTestSummary = Nothing
    Set dctResult = Nothing
    Set dctResultExpected = Nothing
    TestFilesRemove
    Set cllTestFiles = Nothing
    Set Log = Nothing
    
End Sub

Public Property Get ModeRegression() As Boolean:    ModeRegression = bModeRegression:   End Property

Public Property Let ModeRegression(ByVal b As Boolean)
    
    bModeRegression = b
#If mErH Then
    mErH.Regression = b
#End If

End Property

Public Property Let RegrTestSummaryLog(ByVal s As String)
    
    If sRegrTestSummaryLog <> vbNullString Then
        If FSo.FileExists(sRegrTestSummaryLog) Then
            FSo.DeleteFile sRegrTestSummaryLog
        End If
    End If
    sRegrTestSummaryLog = s
    FSo.CreateTextFile s
    
End Property

Public Property Get Result() As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered test result.
' ------------------------------------------------------------------------------
    Dim sId As String:  sId = sTestNumber & "." & dctResult.Count
    
    If VBA.IsObject(dctResult(sId)) _
    Then Set Result = dctResult(sId) _
    Else Result = dctResult(sId)

End Property

Public Property Let Result(ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Register in dctResult a new test result. When the corresponding expected
' result has also already been registered, provide the final test result.
' ------------------------------------------------------------------------------
    Dim sId As String
    sId = TestId(dctResult)
    dctResult.Add sId, r_var
    If dctResultExpected.Exists(sId) _
    Then ResultVersusResultExpectedEvaluation , , sId
    
End Property

Public Property Get ResultExpected() As Variant
' ------------------------------------------------------------------------------
' Returns the most current registered expected test result.
' ------------------------------------------------------------------------------
    Dim sId As String
    sId = TestId(dctResultExpected)
    
    If VBA.IsObject(dctResultExpected(sId)) _
    Then Set ResultExpected = dctResultExpected(sId) _
    Else ResultExpected = dctResultExpected(sId)

End Property

Public Sub EndOfTest()

    If Not ModeRegression Then

#If clsTrc = 1 Then
        If FSo.FileExists(Trc.FileFullName) Then Trc.Dsply
#End If
#If mTrc = 1 Then
        If FSo.FileExists(mTrc.FileFullName) Then mTrc.Dsply
#End If
    End If
    TestFilesRemove

End Sub

Public Property Let ResultExpected(ByVal r_var As Variant)
' ------------------------------------------------------------------------------
' Register in dctResultExpected a new expected test result. When the
' corresponding test result has also already been registered, the final test
' result is provided.
' ------------------------------------------------------------------------------
    Dim sId As String
    
    sId = TestId(dctResultExpected)
    dctResultExpected.Add sId, r_var
    If dctResult.Exists(sId) _
    Then ResultVersusResultExpectedEvaluation , , sId
    
End Property

Public Property Let TestDscrpt(ByVal s As String):  sTestDscrpt = s:                    End Property

Public Property Let TestedComp(ByVal s As String):  sTestedComp = s:                    End Property

Public Property Get TestedComp() As String:         TestedComp = sTestedComp:           End Property

Public Property Let TestedProc(ByVal s As String):  sTestedProc = s
                                                    lMaxLenTestProc = _
                                                    Max(lMaxLenTestProc, Len(s))
End Property

Public Property Get TestedType() As String:         TestedType = sTestedType:           End Property

Public Property Let TestedType(ByVal s As String):  sTestedType = s:                    End Property

Public Property Get TestFile(Optional ByVal t_test_no As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
' - When a test-number (t_test_no) is provided the full default name of an
'   existing test file is returned. When no file with such a test number exists
'   an error is raised.
' - When no test-number is provided, the name of the last provided test file
'   (sTestFile) is returned.
' ------------------------------------------------------------------------------
     Const PROC = "TestFile-Get"
     
     Dim s As String
     
     If t_test_no = vbNullString Then
        If sTestFile = vbNullString _
        Then TestFile = TempFile(sTestFolder) _
        Else TestFile = sTestFile
    Else
        s = TestFileName(t_test_no)
        If FSo.FileExists(s) _
        Then TestFile = FSo.GetFile(s).Path _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "A test-file for the given test-number does not exist in the test folder (" & TestFolder & ")!"
    End If
    
End Property

Public Property Let TestFile(Optional ByVal t_test_no As String = vbNullString, _
                                      ByVal t_file As Variant)
' ------------------------------------------------------------------------------
' - When no test-number (t_test_no) is provided and the provided test-file
'   (t_file) exists, a record is kept in a Collection (cllTestFiles) for the
'   file being removed explicitely with TestFilesRemove or when the class
'   terminates
' - When a test-number is provided the full default name of an existing test
'   file is the current test file
' ------------------------------------------------------------------------------
    
    If TypeName(t_file) = "File" Then t_file = t_file.Path
    If t_test_no = vbNullString Then
        If FSo.FileExists(t_file) _
        Then cllTestFiles.Add t_file
    Else
        sTestFile = TestFileName(t_test_no)
    End If
    
End Property

Public Property Get TestFolder() As String:         TestFolder = sTestFolder:           End Property

Public Property Let TestFolder(ByVal s As String)
    
    If sTestFolder <> vbNullString Then
        If sTestFolder <> s Then
            If FSo.FolderExists(s) Then FSo.DeleteFolder s
        End If
    End If
    sTestFolder = s
    If Not FSo.FolderExists(sTestFolder) _
    Then FSo.CreateFolder sTestFolder
    ChDir sTestFolder

End Property

Public Property Get TestFullName() As String
    Dim s As String
    Select Case Left(sTestedType, 1)
        Case "P":   s = "Property"
        Case "F":   s = "Function"
        Case "S":   s = "Sub     "
        Case "M":   s = "Method  "
    End Select
    TestFullName = sTestNumber & "." & sSubTestId & " " & s & " " & sTestedComp & "." & sTestedProc & ": " & sTestDscrpt

End Property

Public Property Get TestNumber() As String
' ------------------------------------------------------------------------------
' Essential, obligatory distinct identifcation of a test, provided as the very
' first Property for each test.
' ------------------------------------------------------------------------------
    TestNumber = sTestNumber
End Property

Public Property Let TestNumber(ByVal s As String)
' ------------------------------------------------------------------------------
' Essential, obligatory distinct identifcation of a test, provided as the very
' first Property for each test.
' ------------------------------------------------------------------------------
    If sTestNumber <> s Then
        TestReset ' new number indicates new test
        sTestNumber = s
    End If
    
End Property

Public Property Get TestResult() As String: TestResult = sResultEvaluated:   End Property

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

Private Sub ArrayAdd(ByRef a_arr As Variant, _
                     ByVal a_str As String)
' ----------------------------------------------------------------------------
' Adds an item to an array (a_arr) which might not yet is allocated.
' ----------------------------------------------------------------------------
                     
    On Error Resume Next
    ReDim Preserve a_arr(UBound(a_arr) + 1)
    If Err.Number <> 0 Then ReDim a_arr(0)
    a_arr(UBound(a_arr)) = a_str
    
End Sub

Public Function ArrayAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In v_items
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function ArrayAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In v_items
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

Public Function ArrayAsFile(ByVal v_items As Variant, _
                   Optional ByRef v_file_name As String = vbNullString, _
                   Optional ByVal v_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
      
    If v_file_name = vbNullString Then v_file_name = TempFile
    StringAsFile ArrayAsString(v_items), v_file_name, v_file_append
    Set ArrayAsFile = FSo.GetFile(v_file_name)
    
End Function

Public Function ArrayAsString(ByVal a_array As Variant, _
                     Optional ByVal a_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (a_array) as string with the items delimited (a_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(a_array, a_delim)
End Function

Public Function ArrayIsAllocated(ByVal arry As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = UBound(arry) >= LBound(arry)
    On Error GoTo -1
    
End Function

Public Sub AssertedErrors(ParamArray a_app_errs() As Variant)
' ----------------------------------------------------------------------------
' When mErH is not installed (Cond. Comp. Arg. `mErh=1`) this is ignored.
' ----------------------------------------------------------------------------
#If mErH Then
    Const PROC = "AssertedErrors"

    On Error GoTo xt
    If ArrayIsAllocated(a_app_errs) Then
        Select Case UBound(a_app_errs)
            Case 0: mErH.Asserted a_app_errs(0)
            Case 1: mErH.Asserted a_app_errs(0), a_app_errs(1)
            Case 2: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2)
            Case 3: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3)
            Case 4: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4)
            Case 5: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5)
            Case 6: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6)
            Case 7: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6), a_app_errs(7)
            Case 8: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(5), a_app_errs(6), a_app_errs(7), a_app_errs(8)
            Case 9: mErH.Asserted a_app_errs(0), a_app_errs(1), a_app_errs(2), a_app_errs(3), a_app_errs(4), a_app_errs(6), a_app_errs(6), a_app_errs(7), a_app_errs(8), a_app_errs(9)
            Case Else
                Err.Raise AppErr(1), ErrSrc(PROC), "Max number of assertable errors (10) exceeded!"
        End Select
    End If

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
#End If
End Sub

Public Sub BoTP()
' ------------------------------------------------------------------------------
' Begin of Testesd Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP Me.TestFullName
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP Me.TestFullName
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP Me.TestFullName
#End If
End Sub

Public Function CollectionAsArray(ByVal v_items As Collection) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In v_items
        If IsObject(v, s) _
        Then ArrayAdd arr, s _
        Else ArrayAdd arr, v
    Next v
    CollectionAsArray = arr
    
End Function

Public Function CollectionAsDictionary(ByVal v_items As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In v_items
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

Public Function CollectionAsFile(ByVal c_items As Collection, _
                        Optional ByRef c_file As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Transfers the items of a Collection (c_coll) as records/lines to a
' file (c_file_name), optionally appended (c_file_append).
' ----------------------------------------------------------------------------

    If c_file = vbNullString Then c_file = TempFile
    StringAsFile CollectionAsString(c_items), c_file, c_file_append
    Set CollectionAsFile = FSo.GetFile(c_file)

End Function

Public Function CollectionAsString(ByVal c_coll As Collection, _
                          Optional ByVal c_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    Dim sName   As String
    Dim v       As Variant
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then s = s & sDelim & sName _
        Else s = s & sDelim & v
        sDelim = c_delimiter
    Next v
    CollectionAsString = s

End Function

Private Function IsObject(ByVal i_var As Variant, _
                          ByRef i_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
' is an object. When the object does not have a Name property an error is
' raised.
' ----------------------------------------------------------------------------
    Const PROC = "IsObject"
    
    If Not VBA.IsObject(i_var) Then Exit Function
    IsObject = True
    On Error Resume Next
    i_name = i_var.Name
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "VarTrans tries to use the Name property of an object when it is to be " & _
         "transferred into a string which is the case when String, Array, or File " & _
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
End Function

Public Function DictionaryAsArray(ByVal v_items As Dictionary) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In v_items
        If IsObject(v, s) _
        Then ArrayAdd arr, s _
        Else ArrayAdd arr, v
    Next v
    DictionaryAsArray = arr
    
End Function

Public Function ObjectAsString(ByVal o As Object) As String
    On Error Resume Next
    ObjectAsString = o.Name
    If Err.Number <> 0 Then ObjectAsString = "Object_" & o.HashCode
End Function

Public Function DictionaryAsCollection(ByVal v_items As Dictionary) As Collection
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In v_items
            .Add v
        Next v
    End With
    Set DictionaryAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function DictionaryAsDictionary(ByVal v_items As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns the Dictionary (v_items) with the keys in ascending order.
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    With dct
        For Each v In v_items
            .Add v, vbNullString
        Next v
    End With
    Set DictionaryAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function DictionaryAsFile(ByVal v_items As Dictionary, _
                        Optional ByRef v_file_name As String = vbNullString, _
                        Optional ByVal v_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    If v_file_name = vbNullString Then v_file_name = TempFile
    StringAsFile DictionaryAsString(v_items), v_file_name, v_file_append
    Set DictionaryAsFile = FSo.GetFile(v_file_name)

End Function

Public Function BooleanAsString(ByVal b As Boolean) As String
    If b Then BooleanAsString = "TRUE" Else BooleanAsString = "FALSE"
End Function

Public Function DictionaryAsString(ByVal v_items As Dictionary) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    
    For Each v In v_items
        Select Case True
            Case VBA.IsObject(v):           s = s & sDelim & ObjectAsString(v)
            Case VarType(v) = vbBoolean:    s = s & sDelim & BooleanAsString(v)
            Case Else:                      s = s & sDelim & CStr(v)
        End Select
        sDelim = vbCrLf
    Next v
    DictionaryAsString = s

End Function

Public Sub DsplySummary()
    Const PROC = "DsplySummary"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lMax        As Long
    Dim lMaxDesc    As Long
    Dim s           As String
    Dim arr         As Variant
    Dim cll         As Collection
    Dim sType       As String
    Dim bMoreComps  As Boolean
    
    If Not bModeRegression Then Exit Sub ' is bypassed except ModeRegression = True
    
    s = "Summary Regression-Test for component """ & sTestedComp & """ on " & Format(Now(), "YYYY.MM.DD") & " at " & Format(Now(), "hh.mm.ss")
    If FSo.GetFile(sRegrTestSummaryLog).Size > 1 Then
        arr = Split(FileAsString(sRegrTestSummaryLog), vbCrLf)
        VarAsFile String(Len(arr(UBound(arr) - 1)), "="), sRegrTestSummaryLog, True
    End If
    StringAsFile s, sRegrTestSummaryLog, True
    StringAsFile String(Len(s), "-"), sRegrTestSummaryLog, True
    For Each v In dctTestSummary
        Set cll = dctTestSummary(v)
'       1 Test-Id
'       2 sTestedType
'       3 sTestedComp
'       4 sTestedProc
'       5 sTestDscrpt
'       6 sResultEvaluated
        If sTestedComp <> vbNullString And sTestedComp <> cll(3) Then
            bMoreComps = True
        ElseIf sTestedComp = vbNullString Then sTestedComp = cll(3)
        End If

        If bMoreComps _
        Then lMax = Max(lMax, Len(cll(3) & "." & cll(4)) + 2) _
        Else lMax = Max(lMax, Len(cll(4)) + 2)
        lMaxDesc = Max(lMaxDesc, Len(cll(5)) + 2)
    Next v
    
    For Each v In dctTestSummary
        Set cll = dctTestSummary(v)
        
        Select Case Left(cll(2), 1)
            Case "P":   sType = "Property"
            Case "F":   sType = "Function"
            Case "S":   sType = "Sub     "
            Case "M":   sType = "Method  "
        End Select
        
        If bMoreComps _
        Then s = cll(3) & "." & cll(4) _
        Else s = cll(4)
        
        StringAsFile s_strng:=cll(1) & " " & _
                              sType & " " & _
                              s & " " & _
                              String(lMax - Len(s), ".") & _
                              " " & _
                              cll(5) & _
                              String(lMaxDesc - Len(cll(5)) + 5, ".") & _
                              ": " & cll(6) _
                   , s_file:=sRegrTestSummaryLog _
                   , s_file_append:=True
    Next v
    ShellRun sRegrTestSummaryLog
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub EoTP()
' ------------------------------------------------------------------------------
' End of Tested Procedure: Interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP Me.TestFullName, ">" & Me.TestResult
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP Me.TestFullName, ">" & Me.TestResult
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP Me.TestFullName, ">" & Me.TestResult
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
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
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
    If err_source = vbNullString Then err_source = Err.Source
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
            Else: ErrType = "VB Runtime Error "
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
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsTestAid." & sProc
End Function

Public Function ExpectedTestResultFile(ByVal e_test_no As String, _
                              Optional ByVal e_test_folder As String = vbNullString) As File
' ----------------------------------------------------------------------------
' When no test-folder (e_test_folder) is provided, the function returns a
' default expected-test-result-file as object, else one with a default name
' in the provided folder.
' ----------------------------------------------------------------------------
     Const PROC = "ExpectedTestResultFile"
     
     Dim sFile As String
     
     If e_test_folder = vbNullString _
     Then e_test_folder = sTestFolder _
     Else If Not FSo.FolderExists(e_test_folder) _
     Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided folder """ & e_test_folder & """ does not exist!"
     
     sFile = ExpectedTestResultFileName(e_test_no, e_test_folder)
     With FSo
        If .FileExists(sFile) _
        Then Set ExpectedTestResultFile = FSo.GetFile(sFile) _
        Else Err.Raise AppErr(1), ErrSrc(PROC), "An expected test result file named """ & sFile & """ does not exist!"
    End With
     
End Function

Private Function ExpectedTestResultFileName(ByVal e_test_no As String, _
                                   Optional ByVal e_test_folder As String = vbNullString) As String
' ----------------------------------------------------------------------------
' When no test-folder (e_test_folder) is provided, the function returns a
' default expected-test-result-file, else a default name in the provided
' folder.
' ----------------------------------------------------------------------------
     If e_test_folder = vbNullString Then e_test_folder = sTestFolder
    ExpectedTestResultFileName = e_test_folder & "\ResultExpected-" & e_test_no & ".dat"

End Function

Public Function FileAsArray(ByVal f_file As Variant, _
                   Optional ByVal f_empty_excluded) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    For Each v In Split(FileAsString(f_file, vbCrLf), vbCrLf)
        If f_empty_excluded Then
            If Trim$(v) <> vbNullString Then ArrayAdd arr, v
        Else
            ArrayAdd arr, v
        End If
    Next v
    FileAsArray = arr
    
End Function

Public Function FileAsCollection(ByVal v_items As File) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll As New Collection
    With cll
        For Each v In Split(FileAsString(v_items.Path, vbCrLf), vbCrLf)
            .Add v
        Next v
    End With
    Set FileAsCollection = cll
    Set cll = Nothing

xt: Exit Function
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FileAsDictionary(ByVal v_items As File) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the file's records (vulgo text lines) are returned as
'            Directory keys, lines will be unified. I e. line will exist only
'            once. To make this restriction productive, the number of
'            occurrences of each line is returned as item.
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim s   As String
    
    With dct
        For Each v In Split(FileAsString(v_items.Path, vbCrLf), vbCrLf)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function FileAsString(Optional ByVal f_file_full_name As String, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    On Error Resume Next
    Open f_file_full_name For Input As #1
    If Err.Number <> 0 Then GoTo xt
    FileAsString = Input$(lOf(1), 1)
    Close #1
    
    Select Case True
        Case InStr(FileAsString, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(FileAsString, vbCr) <> 0:   f_split = vbCr
        Case InStr(FileAsString, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(FileAsString, 1) = VBA.Chr(26) Then
        FileAsString = Left(FileAsString, Len(FileAsString) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    Do While Right(FileAsString, Len(f_split)) = f_split
        FileAsString = Left(FileAsString, Len(FileAsString) - Len(f_split))
        If Len(FileAsString) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        FileAsString = FileAsStringEmptyExcluded(FileAsString)
    End If
    
xt:
End Function

Private Function FileAsStringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileAsStringTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = s_s
    
End Function

Private Function FileAsStringTrimmed(ByVal s_s As String, _
                            Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    
    s = s_s
    '~~ Eliminate any leading empty items
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    '~~ Eliminate any trailing empty items
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    FileAsStringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

'Private Function IsObject(ByVal i_var As Variant, _
'                          ByRef i_name As String) As Boolean
'' ----------------------------------------------------------------------------
'' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
'' is an object. When the object does not have a Name property an error is
'' raised.
'' ----------------------------------------------------------------------------
'    Const PROC = "IsObject"
'
'    If Not VBA.IsObject(i_var) Then Exit Function
'    IsObject = True
'    On Error Resume Next
'    i_name = i_var.Name
'    If Err.Number <> 0 _
'    Then Err.Raise AppErr(1), ErrSrc(PROC), _
'         "A function tried to use the Name property of an object when it is to be " & _
'         "transferred into a string which is the case when String, Array, or File " & _
'         "is the target format. However, the current item is an object which does " & _
'         "not have a Name property!"
'
'End Function

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

Private Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ------------------------------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Sub ResultEvaluated(Optional ByVal r_result As Collection = Nothing, _
                            Optional ByVal r_result_expected As Collection = Nothing, _
                            Optional ByVal r_test_id As String = vbNullString)
' ------------------------------------------------------------------------------
' When test collections (r_result and r_result_expected are provided these
' collections are compared else the Result and the ResultExpected identified
' by the test-number in the dctResult and dctResultExpected Dictionaries.
' The first choice allows basic without using any function tested elsewhere.
' ------------------------------------------------------------------------------
    Const PROC = "ResultEvaluated"
    
    On Error GoTo eh
    Dim i As Long
    
    sResultEvaluated = RESULT_PASSED ' the default
    If r_result Is Nothing And r_result_expected Is Nothing Then
        If StrComp(VarAsString(dctResult(r_test_id)) _
                 , VarAsString(dctResultExpected(r_test_id)) _
                 , vbBinaryCompare) <> 0 _
        Then sResultEvaluated = RESULT_FAILED
    Else
        If r_result.Count <> r_result_expected.Count Then
            sResultEvaluated = RESULT_FAILED
            GoTo xt
        Else
            For i = 1 To r_result.Count
                If StrComp(CStr(r_result(i)), CStr(r_result_expected(i)), vbBinaryCompare) <> 0 Then
                    sResultEvaluated = RESULT_FAILED
                    GoTo xt
                End If
            Next i
        End If
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ResultFailedDetailsLog(Optional ByVal r_result As Collection = Nothing, _
                                   Optional ByVal r_result_expected As Collection = Nothing, _
                                   Optional ByVal r_test_id As String = vbNullString)
' ------------------------------------------------------------------------------
' Displays the result of a test versus its expected result. The result is
' presented in three columns: Passed/Failed Result Result Expected.
' Called by ResultVersusResultExpectedEvaluation.
' ------------------------------------------------------------------------------
    Const PROC = "ResultFailedDetailsLog"
    
    On Error GoTo eh
    Dim cllResult           As New Collection
    Dim cllResultExpected   As New Collection
    Dim i                   As Long
    Dim lMaxResult          As Long
    Dim lMaxResultExpected  As Long
    Dim v                   As Variant
    
    If Not r_result Is Nothing And Not r_result_expected Is Nothing Then
        Set cllResult = r_result
        Set cllResultExpected = r_result_expected
    Else
        Set cllResult = VarAsCollection(dctResult(r_test_id))
        Set cllResultExpected = VarAsCollection(dctResultExpected(r_test_id))
    End If
    
    For Each v In cllResult
        lMaxResult = Max(lMaxResult, Len(v) + 2)
    Next v
    For Each v In cllResultExpected
        lMaxResultExpected = Max(lMaxResultExpected, Len(v) + 2)
    Next v

    With Log
        .Title "Test Result details: " & TestFullName
        .FileFullName = sTestFolder & "\Result_" & r_test_id & ".log"
        .NewFile
        .AlignmentHeaders "C", "L", "L"
        .MaxItemLengths 8, lMaxResult, lMaxResultExpected
        .Headers "P/F", "Result", "Result Expected"
        .AlignmentItems "C", "L", "L"
        For i = 1 To Min(cllResult.Count, cllResultExpected.Count)
            If StrComp(CStr(cllResult(i)), CStr(cllResultExpected(i)), vbBinaryCompare) = 0 Then
                .Entry "Passed", cllResult(i), cllResultExpected(i)
            Else
                .Entry "Failed", cllResult(i), cllResultExpected(i)
            End If
        Next i
        If cllResult.Count > cllResultExpected.Count Then
            For i = i To cllResult.Count
                .Entry "Failed", CStr(cllResult(i)), vbNullString
            Next i
        ElseIf cllResultExpected.Count > cllResult.Count Then
            For i = i To cllResultExpected.Count
                .Entry "Failed", vbNullString, CStr(cllResultExpected(i))
            Next i
        End If
        .Dsply
    End With

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ResultVersusResultExpectedEvaluation(Optional ByVal r_result As Collection = Nothing, _
                                                Optional ByVal r_result_expected As Collection = Nothing, _
                                                Optional ByVal r_test_id As String = vbNullString)
' ----------------------------------------------------------------------------
' When a result Collection (r_result) and a result expected Collection
' (r_result_expected) is provided the two peovided Collections are compared,
' else the Result and the ResultExpected identified by the test-number in the
' dctResult and dctResultExpected Dictionaries.
' In regression mode the result is provided as summary only, else a detailed
' result is only displayed when the test has failed.
' ----------------------------------------------------------------------------
    Const PROC = "ResultVersusResultExpectedEvaluation"
    
    On Error GoTo eh
    Dim cll As Collection
        
    ResultEvaluated r_result, r_result_expected, r_test_id
    If bModeRegression Then
        If Not r_result Is Nothing And Not r_result_expected Is Nothing Then
            '~~ The Result and the ResultExpected has been provided explicitely.
            '~~ For a correct test-id, both are saved to the dctResult and dctResultExpected
            r_test_id = TestId(dctResult)
            dctResult.Add r_test_id, r_result
            dctResultExpected.Add r_test_id, r_result_expected
        End If
        '~~ In regression mode the result is saved for being displayed as summary.
        Set cll = New Collection
        With cll
            .Add r_test_id
            .Add sTestedType
            .Add sTestedComp
            .Add sTestedProc
            .Add sTestDscrpt
            .Add sResultEvaluated
        End With
        dctTestSummary.Add r_test_id, cll
        Set cll = Nothing
    Else
        If sResultEvaluated = RESULT_FAILED Then ResultFailedDetailsLog r_result, r_result_expected, r_test_id
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ShellRun(ByVal sr_string As String, _
                 Optional ByVal sr_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function SplitIndctr(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns the split indicator in a string (s), which is either vbCrLf or the
' alternate "|&|". When none is found the split indicator defaults to vbCrLf.
' ----------------------------------------------------------------------------
    SplitIndctr = vbCrLf ' The dafault
    
    Select Case True
        Case InStr(s, vbCrLf) <> 0: SplitIndctr = vbCrLf
        Case InStr(s, "|&|") <> 0:  SplitIndctr = "|&|"
        Case InStr(s, ",") <> 0:  SplitIndctr = ","
    End Select

End Function

Public Function StringAsArray(ByVal s_items As String, _
                     Optional ByVal s_split As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    StringAsArray = Split(s_items, s_split)

End Function

Public Function StringAsCollection(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    With cll
        For Each v In Split(s_items, s_split)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function StringAsDictionary(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    
    With dct
        For Each v In Split(s_items, s_split)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                l = dct(v) + 1
                .Remove v
                .Add v, l
            End If
        Next v
    End With
    Set StringAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function StringAsFile(ByVal s_strng As String, _
                    Optional ByRef s_file As Variant = vbNullString, _
                    Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Public Function StringAsString(ByVal s_items As String, _
                      Optional ByVal s_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter string replaced with vbCrLf.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)

    StringAsString = Replace(s_items, s_split, vbCrLf)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function TempFile(Optional ByVal f_path As String = vbNullString, _
                         Optional ByVal f_extension As String = ".txt", _
                         Optional ByVal f_create_as_textstream As Boolean = True) As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' The returned temporary file is registered in cllTestFiles for being removed
' either explicitely with TestFilesRemove or implicitly when the class
' terminates.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    If f_create_as_textstream Then FSo.CreateTextFile sTemp
    cllTestFiles.Add sTemp

End Function

Private Function TestFileName(ByVal e_test_no As String) As String
    TestFileName = TestFolder & "\Test-" & e_test_no & ".txt"

End Function

Public Sub TestFilesRemove(ParamArray t_names() As Variant)

    Dim i   As Long
    Dim s   As String
    Dim fle As File
    Dim v   As Variant
    
    
    For Each fle In FSo.GetFolder(TestFolder).Files
        '~~ Remove any temporary test file
        If fle.Name Like "rad*" Then
            FSo.DeleteFile fle.Path
        End If
        '~~ Remove when explicitely provided named
        If ArrayIsAllocated(t_names) Then
            For Each v In t_names
                If fle.Name Like v & "*" Then
                    FSo.DeleteFile fle.Path
                End If
            Next v
        End If
    Next fle
    
    If Not cllTestFiles Is Nothing Then
        If cllTestFiles.Count > 0 Then
            With cllTestFiles
                For i = .Count To 1 Step -1
                    s = cllTestFiles(i)
                    If FSo.FileExists(s) Then
                        Kill s
                    End If
                    .Remove i
                Next i
            End With
        End If
        Set cllTestFiles = Nothing
        Set cllTestFiles = New Collection
    End If
    
End Sub

Private Function TestId(ByVal dct As Dictionary) As String
    TestId = sTestNumber & "." & dct.Count + 1
End Function

Private Sub TestReset()
    
    dctResult.RemoveAll
    dctResultExpected.RemoveAll
    sTestFile = vbNullString

End Sub

Public Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsArray"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case TypeName(v_items) = "File":        VarAsArray = FileAsArray(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsArray = DictionaryAsArray(v_items)
        Case VarType(v_items) = vbArray:        VarAsArray = v_items
        Case VarType(v_items) = vbString:       VarAsArray = StringAsArray(v_items)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Public Function VarAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsCollection = v_items
        Case TypeName(v_items) = "File":        Set VarAsCollection = FileAsCollection(v_items)
'        Case TypeName(v_items) = "Dictionary":  Ser VarAsCollection = DictionaryAsCollection(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsCollection = ArrayAsCollection(v_items)
        Case VarType(v_items) = vbBoolean:      cll.Add BooleanAsString(v_items)
                                                Set VarAsCollection = cll
        Case VarType(v_items) = vbString:       Set VarAsCollection = StringAsCollection(v_items)
        Case Else:                              cll.Add v_items
                                                Set VarAsCollection = cll

    End Select
    
End Function

Public Function FileAsFile(ByVal f_file_in As File, _
                           ByVal f_file_out As String, _
                  Optional ByVal f_append As Boolean = False, _
                  Optional ByVal f_rename As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a file (f_file_in) as a file with another full name.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsFile"
    
    On Error GoTo eh
    With FSo
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsString(f_file_in), f_file_out, True
            Case Else:                          Err.Raise AppErr(2), ErrSrc(PROC), "Rename  a n d  append is not supported!"
        End Select
    
        Set FileAsFile = .GetFile(f_file_out)
    End With
    
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function VarAsFile(ByVal v_items As Variant, _
                  Optional ByVal v_file As String, _
                  Optional ByVal v_append As Boolean = False) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsFile"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsFile = CollectionAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) = "File"
        Case TypeName(v_items) = "Dictionary":  Set VarAsFile = DictionaryAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbArray:        Set VarAsFile = ArrayAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbString:       Set VarAsFile = StringAsFile(v_items, v_file, v_append)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Public Function VarAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsDictionary = CollectionAsDictionary(v_items)
        Case TypeName(v_items) = "File":        Set VarAsDictionary = FileAsDictionary(v_items)
        Case TypeName(v_items) = "Dictionary":  Set VarAsDictionary = DictionaryAsDictionary(v_items) ' sort key ascending
        Case VarType(v_items) = vbArray:        Set VarAsDictionary = ArrayAsDictionary(v_items)
        Case VarType(v_items) = vbBoolean:      dct.Add BooleanAsString(v_items), vbNullString
                                                Set VarAsDictionary = dct
        Case VarType(v_items) = vbString:       dct.Add v_items, vbNullString
                                                Set VarAsDictionary = dct
    End Select
    
End Function

Public Function VarAsString(ByVal v_items As Variant) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------

    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsString = CollectionAsString(v_items)
        Case TypeName(v_items) = "File":        VarAsString = FileAsString(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsString = DictionaryAsString(v_items)
        Case VarType(v_items) = vbArray:        VarAsString = ArrayAsString(v_items)
        Case VarType(v_items) = vbBoolean:      VarAsString = BooleanAsString(v_items)
        Case VarType(v_items) = vbString:       VarAsString = StringAsString(v_items)
        Case Else:                              VarAsString = CStr(v_items)
    End Select
    
End Function

Public Sub CondCompArgSet(ByVal c_arg As String)
    Const PROC = "CondCompArgSet"
    
    On Error GoTo eh
    Dim sArg        As String
    Dim xlApp       As Application
    Dim wbTarget    As Workbook
    Dim cmb         As CommandBar
    Dim lngHWnd     As Long
    Dim lngHDialog  As Long
    Dim lngHEdit    As Long
    Dim lngHButton  As Long
    Dim cbb         As CommandBarButton
    
    sArg = c_arg

    Set xlApp = CreateObject("Excel.Application")
'    xlApp.Visible = False

    Set wbTarget = ThisWorkbook

    'Launch the VBA Project Properties Dialog
    Set cbb = xlApp.VBE.CommandBars(1).FindControl(id:=2578, recursive:=True)
    Debug.Print cbb.Caption
    cbb.Execute
    

    'Get the handle of the "VBAProject" Window
    lngHWnd = FindWindow("#32770", vbNullString)
    If lngHWnd = 0 Then
        MsgBox "VBAProject Property Window not found!"
        GoTo xt
    End If

    'Get the handle of the dialog
    lngHDialog = FindWindowEx(lngHWnd, ByVal 0&, "#32770", vbNullString)
    If lngHDialog = 0 Then
        MsgBox "VBAProject Property Window could not be accessed!"
        GoTo xt
    End If

    'Get the handle of the 5th edit box
    lngHEdit = fctLngGetHandle("Edit", lngHDialog, 5)
    If lngHEdit = 0 Then
        MsgBox "Conditional Compilation Arguments box could not be accessed!"
        GoTo xt
    End If

    'Enter new argument
    SendMessage lngHEdit, WM_SETTEXT, False, ByVal sArg
    DoEvents

    'Get the handle of the second button box (=OK button)
    lngHButton = fctLngGetHandle("Button", lngHWnd)
    If lngHButton = 0 Then
        MsgBox "Could not find OK button!"
        GoTo xt
    End If

    'Click the OK Button
    SendMessage lngHButton, BM_CLICK, 0, vbNullString

xt: xlApp.Visible = True
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function fctLngGetHandle(strClass As String, lngHParent As Long, _
    Optional Nth As Integer = 1) As Long
    Dim lngHandle As Long
    Dim i As Integer

    lngHandle = FindWindowEx(lngHParent, ByVal 0&, strClass, vbNullString)
    If Nth = 1 Then GoTo Finalize

    For i = 2 To Nth
        lngHandle = FindWindowEx(lngHParent, lngHandle, strClass, vbNullString)
    Next
Finalize:
    fctLngGetHandle = lngHandle
End Function


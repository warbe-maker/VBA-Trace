VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Property Get ColMaxFormHeight() As Range:                               Set ColMaxFormHeight = Me.Range("colMaxFormHeight"):                                End Property

Private Property Get ColMaxFormHeightIncrDecr() As Range:                       Set ColMaxFormHeightIncrDecr = Me.Range("colMaxFormHeightIncDecr"):                 End Property

Private Property Get ColMaxFormWidth() As Range:                                Set ColMaxFormWidth = Me.Range("colMaxFormWidth"):                                  End Property

Private Property Get ColMaxFormWidthIncrDecr() As Range:                        Set ColMaxFormWidthIncrDecr = Me.Range("colMaxFormWidthIncDecr"):                   End Property

Private Property Get ColMinFormHeightIncrDecr() As Range:                       Set ColMinFormHeightIncrDecr = Me.Range("colMinFormHeightIncDecr"):                 End Property

Private Property Get ColMinFormWidth() As Range:                                Set ColMinFormWidth = Me.Range("colMinFormWidth"):                                  End Property

Private Property Get ColMinFormWidthIncrDecr() As Range:                        Set ColMinFormWidthIncrDecr = Me.Range("colMinFormWidthIncDecr"):                   End Property

Public Property Get InitMaxFormHeight(Optional ByVal test As Long) As Long:     InitMaxFormHeight = Intersect(RowTest(test), ColMaxFormHeight).value:               End Property

Public Property Get InitMaxFormWidth(Optional ByVal test As Long) As Long:      InitMaxFormWidth = Intersect(RowTest(test), ColMaxFormWidth).value:                 End Property

Public Property Get InitMinFormWidth(Optional ByVal test As Long) As Long:      InitMinFormWidth = Intersect(RowTest(test), ColMinFormWidth).value:                 End Property

Public Property Get MaxFormHeightIncrDecr(Optional ByVal test As Long) As Long: MaxFormHeightIncrDecr = Intersect(RowTest(test), ColMaxFormHeightIncrDecr).value:   End Property

Public Property Get MaxFormWidthIncrDecr(Optional ByVal test As Long) As Long:  MaxFormWidthIncrDecr = Intersect(RowTest(test), ColMaxFormWidthIncrDecr).value:     End Property

Public Property Get MinFormHeightIncrDecr(Optional ByVal test As Long) As Long: MinFormHeightIncrDecr = Intersect(RowTest(test), ColMinFormHeightIncrDecr).value:   End Property

Public Property Get MinFormWidthIncrDecr(Optional ByVal test As Long) As Long:  MinFormWidthIncrDecr = Intersect(RowTest(test), ColMinFormWidthIncrDecr).value:     End Property

Private Property Get RowTest(Optional ByVal test As Long) As Range:             Set RowTest = Me.Range("rowTest" & test):                                           End Property

Public Property Get TestDescription(Optional ByVal test As Long) As String
    With Me
        TestDescription = Intersect(.Range("rowTest" & test), .Range("colTestDescription")).value
    End With
End Property

Public Property Let TestDescription(Optional ByVal test As Long, ByVal s As String)
    With Me
        Intersect(.Range("rowTest" & test), .Range("colTestDescription")).value = s
    End With
End Property

' -------------------------------------------------------------------------------------------------
' Procedures for test start via Command Buttons on Test Worksheet
Private Sub cmdTest1_Click():   mTest.Test_01_WidthDeterminedByMinimumWidth:                End Sub
Private Sub cmdTest2_Click():   mTest.Test_02_WidthDeterminedByTitle:                       End Sub
Private Sub cmdTest3_Click():   mTest.Test_03_WidthDeterminedByMonoSpacedMessageSection:    End Sub
Private Sub cmdTest4_Click():   mTest.Test_04_WidthDeterminedByReplyButtons:                End Sub
Private Sub cmdTest5_Click():   mTest.Test_05_MonoSpacedSectionWidthExceedsMaxFormWidth:    End Sub
Private Sub cmdTest6_Click():   mTest.Test_06_MonoSpacedMessageSectionExceedMaxFormHeight:  End Sub
Private Sub cmdTest7_Click():   mTest.Test_07_AllInOne:                                     End Sub
Private Sub cmdTest8_Click():   mTest.Test_08_MostlyButtons:                                End Sub
Private Sub cmdTest9_Click():   mTest.Test_09_ButtonsMatrix:                                End Sub
' -------------------------------------------------------------------------------------------------


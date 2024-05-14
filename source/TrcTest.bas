Attribute VB_Name = "TrcTest"
Option Explicit

Public Trc      As clsTrc
Public TestAid  As New clsTestAid

Public Sub Prepare()
    
    Dim s As String

    s = TestAid.TestFolder & "\TestExecTrace.log"
    If Not TestAid.ModeRegression Then
#If clsTrc Then
        Set Trc = New clsTrc
        Trc.NewFile s
#ElseIf mTrc = 1 Then
        With New FileSystemObject
            If .FileExists(s) Then .DeleteFile s
        End With
        mTrc.NewFile s
#End If
    End If
    
End Sub


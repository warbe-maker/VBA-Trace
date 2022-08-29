VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbTrcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' -----------------------------------------------------------------------
' Workbook Trc.xlsm
'          Development and test Workbook for the Common Component 'mTrc'.
'
' See https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
'
' W. Rauschenberger, Berlin, June 2020
' -----------------------------------------------------------------------
Private Const HOSTED_RAWS = "mTrc"

Private Sub Workbook_Open()
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub


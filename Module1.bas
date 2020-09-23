Attribute VB_Name = "Module1"
'- Add to Reference
'1- Crystal Reports ActiveX Designer Run Time Library 10
'2- Crystal ActiveX Report Viewer Library 10
'COMPATIBLE VERSION = CRYSTAL REPORT 10

 Public CRReport As New CRAXDRT.Report
 Public CRApp As New CRAXDRT.Application
 
 Public Sub CrysRpt(CrytalOCX As Object, RptPath As String, DBPath As String, DBPassword As String)

  '-Assign Path to Report
  Set CRReport = CRApp.OpenReport(App.Path & RptPath)
 
  On Error GoTo ExitLabel
  
  With CRReport '-// Log-On to Database with Password Protect

  '-//====================================
  .Database.Tables(1).SetLogOnInfo App.Path & DBPath, App.Path & DBPath, "", DBPassword
 
  End With

    With CrytalOCX

        'Set the source of Report viewer to the path of Report you made
        .ReportSource = CRReport

        'View/Display Report
        .ViewReport

    End With

    CRApp.CanClose

Exit Sub

ExitLabel:     MsgBox Err.Number & " " & Err.Description, vbCritical

End Sub

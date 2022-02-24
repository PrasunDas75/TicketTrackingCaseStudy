VERSION 5.00
Begin VB.MDIForm MDIHome 
   BackColor       =   &H8000000C&
   Caption         =   "Home"
   ClientHeight    =   8865
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14715
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuCreate 
      Caption         =   "&Create Ticket"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View Ticket"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close Ticket"
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&View Report"
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "&Log Out"
   End
End
Attribute VB_Name = "MDIHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
If ChkDevops = True Then
    mnuCreate.Visible = False
Else
    mnuView.Visible = False
    mnuClose.Visible = False
End If
End Sub

Private Sub mnuClose_Click()
frmClosetckts.Show
End Sub

Private Sub mnuCreate_Click()
frmCreateTicket.Show
End Sub

Private Sub mnuLogout_Click()
frmLogin.Show
ChkDevops = False
Unload Me
End Sub

Private Sub mnuView_Click()
frmViewtckts.Show
End Sub

Private Sub mnuReport_Click()
Dim crApp As New CRAXDRT.Application
Dim crRpt As New CRAXDRT.Report

Dim filepath As String

filepath = "F:\0_download\0Eurofins_Training\Case Study\ViewTickets.rpt"

Set crRpt = crApp.OpenReport(filepath)
frmReport.CRViewer91.ReportSource = crRpt
frmReport.CRViewer91.ViewReport
frmReport.Show
End Sub


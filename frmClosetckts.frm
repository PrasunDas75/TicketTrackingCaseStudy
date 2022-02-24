VERSION 5.00
Begin VB.Form frmClosetckts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close Tickets"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3000
      TabIndex        =   4
      Top             =   3840
      Width           =   1230
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   1230
   End
   Begin VB.ComboBox ComTicketId 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtResolution 
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ComboBox ComResolvedby 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblResolution 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   2220
   End
   Begin VB.Label lblResolvedBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolved By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label lblTicketID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1440
   End
End
Attribute VB_Name = "frmClosetckts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myconnection As New dbconnection

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSubmit_Click()
Dim msg As String
Dim myTicket As New SubmitTicket

msg = myTicket.Update(CInt(ComTicketId.Text), ComResolvedby.Text, txtResolution.Text)

MsgBox msg
End Sub

Private Sub Form_Load()
 Dim mydept As New dept
  Dim mydept2 As New dept
  
myconnection.SetUpConnection

Dim item() As String
Dim item2() As String

item = mydept.getEmpD()
item2 = mydept2.getTkt()

For i = 0 To UBound(item)
ComResolvedby.AddItem item(i)
Next

For i = 0 To UBound(item2)
ComTicketId.AddItem item2(i)
Next
End Sub

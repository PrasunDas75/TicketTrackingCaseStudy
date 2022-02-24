VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCreateTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Ticket"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
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
   ScaleHeight     =   5430
   ScaleWidth      =   6555
   Begin VB.ComboBox ComSeverity 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy HH:mm"
      Format          =   190971907
      CurrentDate     =   44615
   End
   Begin VB.TextBox txtDesc 
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox ComEmployee 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2655
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
      TabIndex        =   1
      Top             =   4560
      Width           =   1230
   End
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
      TabIndex        =   0
      Top             =   4560
      Width           =   1230
   End
   Begin VB.Label lblTicketDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Date"
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
      Top             =   1560
      Width           =   2220
   End
   Begin VB.Label lblEmployeeId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EmployeeId"
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
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label lblSeverity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Severity"
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
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Description"
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
      TabIndex        =   4
      Top             =   3000
      Width           =   2220
   End
End
Attribute VB_Name = "frmCreateTicket"
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

msg = myTicket.Save(ComEmployee.Text, Format(Now, "General date"), ComSeverity.Text, txtDesc.Text)

MsgBox msg
End Sub


Private Sub Form_Load()
 Dim mydept As New dept
myconnection.SetUpConnection
Dim item() As String
item = mydept.getEmp()

DTPicker1.Value = Format(Now, "General date")
DTPicker1.Enabled = False

For i = 0 To UBound(item)
ComEmployee.AddItem item(i)
Next

ComSeverity.AddItem "Critical"
ComSeverity.AddItem "Major"
ComSeverity.AddItem "Normal"
End Sub

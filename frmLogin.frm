VERSION 5.00
Object = "*\AUserControlPj.vbp"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
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
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin Project3.UserControlPW UserControlPW1 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Text            =   ""
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   1230
   End
   Begin VB.ComboBox ComDept 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtEid 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblDepartment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1470
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1170
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
      TabIndex        =   0
      Top             =   840
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myconnection As New dbconnection
Dim myAuth As New EmployeeAuth
Dim check As Boolean

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSubmit_Click()
Call myAuth.Authenticate(txtEid.Text, UserControlPW1.Text, ComDept.Text)
check = myAuth.IsAuthenticated
If ComDept.Text = "DEVOPS" And check Then
    ChkDevops = True
    MDIHome.Show
    Unload Me
ElseIf ComDept.Text <> "DEVOPS" And check Then
    MDIHome.Show
    Unload Me
Else
    MsgBox "Invalid Login"
End If

End Sub


Private Sub Form_Load()
 Dim mydept As New dept
myconnection.SetUpConnection
Dim item() As String
item = mydept.getDep()

For i = 0 To UBound(item)
ComDept.AddItem item(i)
Next

End Sub


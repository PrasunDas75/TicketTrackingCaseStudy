VERSION 5.00
Begin VB.UserControl UserControlPW 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4830
   ScaleWidth      =   8715
   Begin VB.TextBox txtText1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "UserControlPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public chkPW As Boolean
Public chkSp As Boolean
Public myText As String

Private Sub txtText1_Change()
If Len(txtText1.Text) > 12 Then MsgBox "Should not be more than 12 characters"
End Sub

Private Sub txtText1_LostFocus()
For i = 1 To Len(txtText1.Text)
    If Not IsNumeric(Mid$(txtText1.Text, i, 1)) And LCase$(Mid$(txtText1.Text, i, 1)) <> UCase$(Mid$(txtText1.Text, i, 1)) Then
        If Mid$(txtText1.Text, i, 1) = UCase$(Mid$(txtText1.Text, i, 1)) Then chkPW = True
    End If
Next

For i = 1 To Len(txtText1.Text)
    If Not IsNumeric(Mid$(txtText1.Text, i, 1)) Then
        If LCase$(Mid$(txtText1.Text, i, 1)) = UCase$(Mid$(txtText1.Text, i, 1)) Then chkSp = True
    End If
Next

If Len(txtText1.Text) < 8 Or Not chkPW Or Not chkSp Then
    MsgBox "Invalid password. Password should be combination of Upper character, Special character and length should be at least 8 character"
    txtText1.SetFocus
Else
    chkPW = False
    chkSp = True
End If
End Sub

Private Sub UserControl_Resize()
txtText1.Left = 0
txtText1.Top = 0
txtText1.Width = UserControl.Width
txtText1.Height = UserControl.Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    myText = PropBag.ReadProperty("Text", myText)
    txtText1.Text = myText
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", myText, UserControl.Name)
    txtText1.Text = myText
End Sub

Public Property Let Text(ByVal newValue As String)
    myText = newValue
    Call UserControl.PropertyChanged("Text")
    txtText1.Text = myText
End Property
Public Property Get Text() As String
    Text = txtText1.Text
End Property


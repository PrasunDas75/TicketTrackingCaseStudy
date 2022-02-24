VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public IsAuthenticated As Boolean
Public chkemp As New EmployeeAuth
'Public isdevops As Boolean

Public Function Authenticate(ByVal EID As String, ByVal Password As String, ByVal Dept As String)
Dim myRecSet As ADODB.Recordset
Dim sql As String
IsAuthenticated = False
sql = "select ea.AuthId from EmployeeAuthentication as ea " & _
            "inner join EMPLOYEE as e " & _
            "on e.EID = ea.EID " & _
            "where ea.EID = '" & EID & "' and ea.[Password]='" & Password & "' and e.Dept = '" & Dept & "'"

Set myRecSet = myConnection.Execute(sql)
If Not myRecSet.EOF Then
    IsAuthenticated = True
    
'    If Dept = "DEVOPS" Then
'        isdevops = True
'    End If
End If

End Function

'Public Function Check() As Boolean
'         Check = isdevops
'End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim i As Integer

Private arr() As String
Dim arrSize As Integer

Public Function getDep() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT distinct(e.dept) FROM employee as e")
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("dept").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    
    getDep = arr
    
End Function


Public Function getEmp() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT e.EID FROM employee as e where e.Dept <> 'DEVOPS'")
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("EID").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    
    getEmp = arr
    
End Function

Public Function getEmpD() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT e.EID FROM employee as e where e.Dept = 'DEVOPS'")
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("EID").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    
    getEmpD = arr
    
End Function

Public Function getTkt() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT t.Ticket_Id FROM Tickets as t")
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("Ticket_Id").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    
    getTkt = arr
    
End Function


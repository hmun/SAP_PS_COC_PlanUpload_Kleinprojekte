VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function addMsg(MyEvent As String, MyQ As String, MyZ As String, MyCode As String, _
                MyText As String, myLevel As Integer) As Integer
On Error GoTo AddMsg_Error
'Dim mydb As DAO.Database
'Dim MyMsgTab As DAO.Recordset
'Dim MyMsgTabName As String
'Dim mysql As String
'Dim MyMsgNrTab As DAO.Recordset
'Dim myDebugLevel As Integer
'  myDebugLevel = CInt(GetLParam("MELD", "DebugLevel"))
' If myLevel > myDebugLevel Then
'    Exit Function
'  End If
'  Set mydb = CurrentDb
'  MyMsgTabName = "T_MELDUNGEN"
'  Set MyMsgTab = mydb.OpenRecordset(MyMsgTabName, dbOpenDynaset)
'  MyMsgTab.AddNew
'  MyMsgTab("EVENT") = Left(MyEvent, 20)
'  MyMsgTab("QUELLE") = Left(MyQ, 10)
'  MyMsgTab("ZIEL") = Left(MyZ, 10)
'  MyMsgTab("ZEIT") = Now()
'  MyMsgTab("CODE") = Left(MyCode, 10)
'  MyMsgTab("TEXT") = Left(MyText, 255)
'  MyMsgTab.Update
  addMsg = True
  Exit Function
AddMsg_Error:
  addMsg = False
  MsgBox "Fehler: " & err.Number & ", " & err.Description, , "AddMsg"
  Exit Function
End Function


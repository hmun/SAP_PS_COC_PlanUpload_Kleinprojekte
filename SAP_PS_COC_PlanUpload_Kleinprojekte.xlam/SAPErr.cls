VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private MyMSG As New TaskMSG

Public Function MSGOnlyProt(MyEvent As String, MyQ As String, MyZ As String, MyCode As String, MyText As String)
On Error GoTo MSGProt_Error
  MyMSG.addMsg MyEvent, MyQ, MyZ, MyCode, MyText, 0
  Exit Function
MSGProt_Error:
  Exit Function
End Function

Public Function MSGProt(MyEvent As String, MyQ As String, MyZ As String, MyCode As String, MyText As String)
On Error GoTo MSGProt_Error
  MyMSG.addMsg MyEvent, MyQ, MyZ, MyCode, MyText, 0
  MSGNonProt MyEvent, MyQ, MyZ, MyCode, MyText
  Exit Function
MSGProt_Error:
  Exit Function
End Function

Public Function MSGNonProt(MyEvent As String, MyQ As String, MyZ As String, MyCode As String, MyText As String)
On Error GoTo MSGNonProt_Error
  MsgBox "Fehler bei Ereignis: " & MyEvent & _
         ", Quelle: " & MyQ & _
         ", Ziel: " & MyZ & _
         ", Fehlernummer: " & MyCode & _
         ", Fehler: " & MyText _
         , , "SAPErr"
  Exit Function
MSGNonProt_Error:
  Exit Function
End Function


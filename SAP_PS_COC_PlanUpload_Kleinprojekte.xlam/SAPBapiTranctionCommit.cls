VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPBapiTranctionCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyFunction As Object
Private MyTranstable As Object
Public IsConnected As Integer
Public RetText As String

Public Function commit() As Integer
Dim ret As Integer
Dim aFunction As Object
On Error GoTo commit_Error
  Set aFunction = MySAPCon.SAPFunctions.add("BAPI_TRANSACTION_COMMIT")
'  Set aRet = aFunction.Tables("TRANSDATA")
  ret = aFunction.call
  commit = ret
  Exit Function
commit_Error:
  commit = False
  MySAPErr.MSGProt "SAPBapiTranctionCommit", "commit", "", err.Number, err.Description
  Exit Function
End Function

Private Sub Class_Initialize()
  IsConnected = SAPCheck()
End Sub

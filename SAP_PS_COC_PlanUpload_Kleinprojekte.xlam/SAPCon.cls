VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public SAPCon As Object
Public SAPFunctions As Object
Private SAPLogon As Object
Public SAPBapiControl As Object

Public Function Init() As Integer
On Error GoTo Init_Error
    Set SAPLogon = CreateObject("SAP.LogonControl.1")
    Set SAPCon = SAPLogon.NewConnection
    Set SAPFunctions = CreateObject("SAP.Functions")
    Set SAPFunctions.Connection = SAPCon
    Set SAPBapiControl = CreateObject("SAP.BAPI.1")
    SAPBapiControl.Connection = SAPCon
    Init = True
    Exit Function
Init_Error:
    MySAPErr.MSGNonProt "SAPCon", "Init", "", err.Number, err.Description
    Init = False
    Exit Function
End Function

Public Function Destruct()
On Error GoTo Destruct_Error
    If Not SAPFunctions Is Nothing Then
        Set SAPFunctions.Connection = Nothing
    End If
    Set SAPFunctions = Nothing
    Set SAPBapiControl = Nothing
    Set SAPCon = Nothing
    Set SAPLogon = Nothing
    Exit Function
Destruct_Error:
    MySAPErr.MSGNonProt "SAPCon", "Destruct", "", err.Number, err.Description
    Destruct = False
    Exit Function
End Function


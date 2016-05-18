Attribute VB_Name = "SAPGlobal"
Public MySAPCon As New SAPCon
Public MySAPErr As New SAPErr

Function SAPInit() As Integer
If MySAPCon.Init Then
  If MySAPCon.SAPCon.Logon Then
    SAPInit = True
  Else
    SAPInit = False
  End If
  Exit Function
Else
  SAPInit = False
End If
End Function

Function SAPCheck() As Integer
  If MySAPCon.SAPCon Is Nothing Then
    If Not SAPInit() Then
      SAPCheck = False
      Exit Function
    End If
  End If
  If MySAPCon.SAPCon.IsConnected <> 1 Then
    If Not SAPInit() Then
      SAPCheck = False
      Exit Function
    End If
  End If
  SAPCheck = True
End Function

Function SAPLogoff() As Integer
    If Not MySAPCon Is Nothing Then
        If Not MySAPCon.SAPCon Is Nothing Then
            If MySAPCon.SAPCon.IsConnected = 1 Then
                MySAPCon.SAPCon.Logoff
            End If
        End If
        MySAPCon.Destruct
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer



Private Sub Class_Initialize()
  IsConnected = SAPCheck()
End Sub

Public Function CheckUser(pUSERNAME As String) As String

On Error GoTo post_Error
Dim lSAPFormat As New SAPFormat
Dim boUser As Object
Dim oActivitygroups As Object
Dim oRETURN As Object

'***** Initialisierung Bapi *****

  Set boUser = MySAPCon.SAPBapiControl.GetSAPObject("USER", UCase(Trim(pUSERNAME))) ' Abfrage über Excel Benutzer
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boUser, "GetDetail", "Return")
  Set oActivitygroups = MySAPCon.SAPBapiControl.DimAs(boUser, "GetDetail", "Activitygroups")
  
  'boUser("USERNAME") = MySAPCon.SAPCon.User
  'UserName:=pUSERNAME,

'***** Abfrage Daten über Bapi *****

  Application.Cursor = xlWait
  boUser.GetDetail Activitygroups:=oActivitygroups, _
                          RETURN:=oRETURN
  Application.Cursor = xlDefault
  
'***** Einlesen der abgefragten Daten *****

  Set bRow = oActivitygroups.Rows(1)
    Dim RetRow As Object
    sPlanning = 0


    CheckUser = pUSERNAME
    

    If oActivitygroups.RowCount <> 0 Then
    
        For Each RetRow In oActivitygroups.Rows
            

        
            If RetRow("AGR_NAME") = "ZMS_CO_EXCEL_PLANDATEN" And Now > RetRow("FROM_DAT") And Now < RetRow("to_DAT") Then
            
                sPlanning = 1
                CheckUser = CheckUser & "  " & RetRow("AGR_NAME")
                

            End If
        
        Next RetRow
    End If
    
'***** Abfragen der Returnmeldungen *****
    
  Set RetRow = Nothing
  
    If oRETURN.RowCount <> 0 Then
      For Each RetRow In oRETURN.Rows
        Set aRow = oRETURN.Rows(1)
          If aRow("TYPE") = "S" Then
        '    'Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        '    'aSAPBapiTranctionCommit.commit
                CheckUser = CheckUser & "  " & aRow("MESSAGE")
        '  Else
        '
        '    For Each RetRow In oRETURN.Rows
        '      CheckUser = CheckUser & ";" & RetRow("MESSAGE")
        '    Next RetRow
          End If
      Next RetRow
    End If
    
  Set boUser = Nothing
  Exit Function
post_Error:
  Application.Cursor = xlDefault
  CheckUser = "Error"
  MySAPErr.MSGProt "User", "GetDetail", "", err.Number, err.Description
  Exit Function
End Function


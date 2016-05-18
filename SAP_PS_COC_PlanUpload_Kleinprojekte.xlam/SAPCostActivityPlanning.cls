VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCostActivityPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
  IsConnected = SAPCheck()
End Sub

Public Function PostPrimCostPer(pData As Collection) As String
On Error GoTo PostPrimCostPer_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oPervalue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Coobject")
  Set oPervalue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Pervalue")
  
Dim aSAP_PC_Plan As SAP_PC_Plan
  Set aSAP_PC_Plan = pData(1)
  oHeaderinfo("CO_AREA") = aSAP_PC_Plan.Header("CO_AREA").Value
  oHeaderinfo("FISC_YEAR") = aSAP_PC_Plan.Header("FISC_YEAR").Value
  oHeaderinfo("PERIOD_FROM") = aSAP_PC_Plan.Header("PERIOD_FROM").Value
  oHeaderinfo("PERIOD_TO") = aSAP_PC_Plan.Header("PERIOD_TO").Value
  oHeaderinfo("VERSION") = aSAP_PC_Plan.Header("VERSION").Value
  oHeaderinfo("PLAN_CURRTYPE") = aSAP_PC_Plan.Header("PLAN_CURRTYPE").Value
  Dim aOField As TField
  Dim aVField As TField
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aPValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aDataRow As Collection
  Dim aVal
  lCnt = 0
  For Each aSAP_PC_Plan In pData
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    For Each aOField In aSAP_PC_Plan.Object
        aCOObjRow.Value(aOField.Name) = aOField.Value
    Next
    Set aIndRow = oIndexstructure.Rows.add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aPValRow = oPervalue.Rows.add
    aPValRow.Value("VALUE_INDEX") = lCnt
    aPValRow.Value("COST_ELEM") = aSAP_PC_Plan.CostElement
'   move the values from the data
    For Each aVField In aSAP_PC_Plan.ValVar
      aPValRow.Value(aVField.Name) = aVField.Value
    Next aVField
    For Each aVField In aSAP_PC_Plan.ValFix
      aPValRow.Value(aVField.Name) = aVField.Value
    Next aVField
  Next aSAP_PC_Plan
  boCostActivityPlanning.PostPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        Pervalue:=oPervalue, _
                        RETURN:=oRETURN
  If oRETURN.RowCount = 0 Then
    PostPrimCostPer = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostPrimCostPer = PostPrimCostPer & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostPrimCostPer_Error:
  PostPrimCostPer = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostPrimCostPer", "", err.Number, err.Description
  Exit Function
End Function

'++ HM160413
Public Function PostPrimCostClear(pData As Collection) As String
On Error GoTo PostPrimCostClear_Error
Dim lSAPFormat As New SAPFormat
Dim boCostActivityPlanning As Object
Dim oHeaderinfo As Object
Dim oIndexstructure As Object
Dim oCoobject As Object
Dim oTotValue As Object
Dim oRETURN As Object
  Set boCostActivityPlanning = MySAPCon.SAPBapiControl.GetSAPObject("CostActivityPlanning")
  Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Return")
  Set oHeaderinfo = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Headerinfo")
  Set oIndexstructure = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Indexstructure")
  Set oCoobject = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Coobject")
  Set oTotValue = MySAPCon.SAPBapiControl.DimAs(boCostActivityPlanning, "PostPrimCost", "Totvalue")
  
  Dim aSAP_PC_Plan As SAP_PC_Plan
  Set aSAP_PC_Plan = pData(1)
  oHeaderinfo("CO_AREA") = aSAP_PC_Plan.Header("CO_AREA").Value
  oHeaderinfo("FISC_YEAR") = aSAP_PC_Plan.Header("FISC_YEAR").Value
  oHeaderinfo("PERIOD_FROM") = aSAP_PC_Plan.Header("PERIOD_FROM").Value
  oHeaderinfo("PERIOD_TO") = aSAP_PC_Plan.Header("PERIOD_TO").Value
  oHeaderinfo("VERSION") = aSAP_PC_Plan.Header("VERSION").Value
  oHeaderinfo("PLAN_CURRTYPE") = aSAP_PC_Plan.Header("PLAN_CURRTYPE").Value
  Dim aObjRow As Object
  Dim aIndRow As Object
  Dim aCOObjRow As Object
  Dim aTValRow As Object
  Dim lRow As Object
  Dim lCnt As Integer
  Dim aDataRow As Collection
  Dim aVal
  lCnt = 0
  For Each aSAP_PC_Plan In pData
    lCnt = lCnt + 1
    Set aCOObjRow = oCoobject.Rows.add
    aCOObjRow.Value("OBJECT_INDEX") = lCnt
    For Each aOField In aSAP_PC_Plan.Object
        aCOObjRow.Value(aOField.Name) = aOField.Value
    Next
    Set aIndRow = oIndexstructure.Rows.add
    aIndRow.Value("OBJECT_INDEX") = lCnt
    aIndRow.Value("VALUE_INDEX") = lCnt
    Set aTValRow = oTotValue.Rows.add
    aTValRow.Value("VALUE_INDEX") = lCnt
    aTValRow.Value("COST_ELEM") = aSAP_PC_Plan.CostElement
'   set the values to 0
    aTValRow.Value("FIX_VALUE") = 0
    aTValRow.Value("DIST_KEY_FIX_VAL") = "1"
    aTValRow.Value("VAR_VALUE") = 0
    aTValRow.Value("DIST_KEY_VAR_VAL") = "1"
    aTValRow.Value("FIX_QUAN") = 0
    aTValRow.Value("DIST_KEY_FIX_QUAN") = "1"
    aTValRow.Value("VAR_QUAN") = 0
    aTValRow.Value("DIST_KEY_VAR_QUAN") = "1"
  Next aSAP_PC_Plan
  boCostActivityPlanning.PostPrimCost Headerinfo:=oHeaderinfo, _
                        Indexstructure:=oIndexstructure, _
                        Coobject:=oCoobject, _
                        TotValue:=oTotValue, _
                        RETURN:=oRETURN
  If oRETURN.RowCount = 0 Then
    PostPrimCostClear = "Success"
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    aSAPBapiTranctionCommit.commit
  Else
    Dim RetRow As Object
    For Each RetRow In oRETURN.Rows
      PostPrimCostClear = PostPrimCostClear & ";" & RetRow("MESSAGE")
    Next RetRow
  End If
  Set boCostActivityPlanning = Nothing
  Exit Function
PostPrimCostClear_Error:
  PostPrimCostClear = "Error"
  MySAPErr.MSGProt "SAPCostActivityPlanning", "PostPrimCostClear", "", err.Number, err.Description
  Exit Function
End Function
'++ HM160413





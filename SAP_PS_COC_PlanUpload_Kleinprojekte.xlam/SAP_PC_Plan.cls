VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_PC_Plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Header As Collection
Public Object As Collection
Public ValFix As Collection
Public ValVar As Collection
Public CostElement As String
Private STfield As TField

Private Sub Class_Initialize()
    Set STfield = New TField
    Set Header = New Collection
    Set Object = New Collection
    Set ValFix = New Collection
    Set ValVar = New Collection
End Sub

Public Function create(pCoAre As String, pFiscy As String, pPfrom As String, _
                       pPto As String, pVers As String, pCurt As String) As SAP_PC_Plan
    Dim newSAP_PC_Plan As SAP_PC_Plan
    Set newSAP_PC_Plan = New SAP_PC_Plan
    
    newSAP_PC_Plan.setHValues pCoAre, pFiscy, pPfrom, pPto, pVers, pCurt
    Set create = newSAP_PC_Plan
End Function

Public Function setHValues(pCoAre As String, pFiscy As String, pPfrom As String, _
                           pPto As String, pVers As String, pCurt As String)
    Dim lSAPFormat As New SAPFormat
    Dim lTField As TField
    Set lTField = STfield.create("CO_AREA", pCoAre)
    Header.add lTField, lTField.Name
    Set lTField = STfield.create("FISC_YEAR", pFiscy)
    Header.add lTField, lTField.Name
    Set lTField = STfield.create("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
    Header.add lTField, lTField.Name
    Set lTField = STfield.create("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
    Header.add lTField, lTField.Name
    Set lTField = STfield.create("VERSION", lSAPFormat.unpack(pVers, 3))
    Header.add lTField, lTField.Name
    Set lTField = STfield.create("PLAN_CURRTYPE", pCurt)
    Header.add lTField, lTField.Name
End Function

Public Function setOValues(pObject As String, pValue As String)
    Dim lTField As TField
    Set lTField = STfield.create(pObject, pValue)
    Object.add lTField, lTField.Name
End Function

Public Function addVValue(pName As String, pValue As String)
    Dim lTField As TField
    Set lTField = STfield.create(pName, pValue)
    ValVar.add lTField, lTField.Name
End Function

Public Function addFValue(pName As String, pValue As String)
    Dim lTField As TField
    Set lTField = STfield.create(pName, pValue)
    ValFix.add lTField, lTField.Name
End Function

Private Function contains(col As Collection, Key As Variant, Optional aType As String = "var") As Boolean
Dim obj As Object
Dim var As Variant
On Error GoTo err
    contains = True
    If aType = "obj" Then
        Set obj = col(Key)
    Else
        var = col(Key)
    End If
    Exit Function
err:
    contains = False
End Function

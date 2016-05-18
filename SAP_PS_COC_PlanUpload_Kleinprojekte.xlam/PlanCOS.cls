VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private STfield As TField

Public aPlanCOS As Collection

Private Sub Class_Initialize()
    Set aPlanCOS = New Collection
    Set STfield = New TField
End Sub

Public Function addCOS(p_Year As Integer, p_Elem As String, p_Value As Double)
Dim aPlanCOSRec As PlanCOSRec
Dim aKey As String
Dim aValue As TField
    aKey = CStr(p_Year) & "-" & p_Elem
    If contains(aPlanCOS, aKey, "obj") Then
        Set aPlanCOSRec = aPlanCOS(aKey)
        Set aValue = aPlanCOSRec.aValue
        aValue.add CDbl(p_Value)
        Set aPlanCOSRec.aValue = aValue
    Else
        Set aPlanCOSRec = New PlanCOSRec
        aPlanCOSRec.setValues p_Year, p_Elem, p_Value
        aPlanCOS.add aPlanCOSRec, aKey
    End If
End Function

Public Function addOtherCOS(p_PlanHours As PlanHours, p_PlanCOSPrice As PlanCOSPrice, p_PlanElemMapping As PlanElemMapping)
Dim aPHours As TField
Dim aPCOSPrice As TField
Dim aPElemMapping As TField
Dim aHours As Double
Dim aPrice As Double
Dim aValue As Double
Dim aYear As String
Dim aCostElement As String
Dim aKey As String
    aCostElement = p_PlanElemMapping.getCElem("COS")
    For Each aPHours In p_PlanHours.aPlanHours
        Set aPlanCOSRec = New PlanCOSRec
        aYear = aPHours.Name
        Set aPCOSPrice = p_PlanCOSPrice.aPlanCOSPrices(aYear)
        aHours = CDbl(aPHours.Value)
        aPrice = CDbl(aPCOSPrice.Value)
        aValue = aPrice * aHours
        aKey = aYear & "-" & aCostElement
        Set aPlanCOSRec = New PlanCOSRec
        aPlanCOSRec.setValues aPHours.Name, aCostElement, CStr(aValue)
        aPlanCOS.add aPlanCOSRec, aKey
    Next
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanCOSRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public aYear As TField
Public aElem As TField
Public aValue As TField

Private STfield As TField

Private Sub Class_Initialize()
    Set STfield = New TField
    Set aYear = New TField
    Set aElem = New TField
    Set aValue = New TField
End Sub

Public Function setValues(pYear As Integer, pElem As String, pValue As Double)
    Set aYear = STfield.create("Year", CStr(pYear))
    Set aElem = STfield.create("Elem", CStr(pElem))
    Set aValue = STfield.create("Value", CStr(pValue))
End Function


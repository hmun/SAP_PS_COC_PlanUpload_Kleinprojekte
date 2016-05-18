VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Value As String

Public Function create(pName As String, pValue As String) As TField
    Set newTField = New TField
    newTField.setValues pName, pValue
    Set create = newTField
End Function

Public Function setValues(pName As String, pValue As String)
    Name = pName
    Value = pValue
End Function

Public Function add(p_Val As Double)
    Dim aVal As Double
    aVal = CDbl(Value)
    aVal = aVal + p_Val
    Value = CStr(aVal)
End Function

Public Function concat(p_Val As String, Optional p_Sep As String = "")
    Value = Value & p_Sep & p_Val
End Function


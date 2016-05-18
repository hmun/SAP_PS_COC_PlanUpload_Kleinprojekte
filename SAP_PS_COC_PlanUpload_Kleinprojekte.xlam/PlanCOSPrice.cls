VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanCOSPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private STfield As TField

Public aPlanCOSPrices As Collection

Private Sub Class_Initialize()
    Set aPlanCOSPrices = New Collection
    Set STfield = New TField
End Sub

Public Function addPrice(p_Year As Integer, p_Price As Double)
Dim aTField As TField
    If contains(aPlanCOSPrices, CStr(p_Year), "obj") Then
        Set aTField = aPlanCOSPrices(CStr(p_Year))
        aTField.add CDbl(p_Price)
    Else
        Set aTField = STfield.create(CStr(p_Year), CStr(p_Price))
        aPlanCOSPrices.add aTField, CStr(p_Year)
    End If
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


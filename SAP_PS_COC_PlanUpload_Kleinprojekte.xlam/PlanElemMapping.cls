VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanElemMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private STfield As TField

Public aPlanElemMappings As Collection

Private Sub Class_Initialize()
    Set aPlanElemMappings = New Collection
    Set STfield = New TField
End Sub

Public Function addElem(p_Elem As String, p_CostElem As String)
Dim aTField As TField
    If contains(aPlanElemMappings, p_Elem, "obj") Then
        Set aTField = aPlanElemMappings(p_Elem)
        aTField.setValues p_Elem, p_CostElem
    Else
        Set aTField = STfield.create(p_Elem, p_CostElem)
        aPlanElemMappings.add aTField, p_Elem
    End If
End Function

Public Function getCElem(p_Elem As String) As String
Dim aTField As TField
    If contains(aPlanElemMappings, p_Elem, "obj") Then
        Set aTField = aPlanElemMappings(p_Elem)
        getCElem = aTField.Value
    Else
        getCElem = ""
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


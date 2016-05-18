VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private STfield As TField

Public aMsgCol As Collection

Private Sub Class_Initialize()
    Set aMsgCol = New Collection
    Set STfield = New TField
End Sub

Public Function addMsg(p_Msg As String, p_Key As String, Optional p_Sep As String = "")
Dim aTField As TField
    If contains(aMsgCol, p_Key, "obj") Then
        If p_Msg <> "" Then
            Set aTField = aMsgCol(p_Key)
            aTField.concat p_Msg, p_Sep
        End If
    Else
        Set aTField = STfield.create(p_Key, p_Msg)
        aMsgCol.add aTField, p_Key
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

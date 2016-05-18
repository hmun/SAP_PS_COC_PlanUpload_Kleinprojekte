VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function unpack(val As String, length As Integer) As String
Dim ZeroStr As String
    If IsNumeric(val) Then
        ZeroStr = "000000000000000000000000000000"
        unpack = Left(ZeroStr, length - Len(val)) & val
    Else
        unpack = val
    End If
End Function


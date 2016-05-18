
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ret = delSAPCommandbar
End Sub

Private Sub Workbook_Open()
  ret = addSAPCommandbar
End Sub
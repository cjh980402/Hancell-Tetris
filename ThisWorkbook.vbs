Private Sub Workbook_Open()
  activesheet.protect "Tkdlqjrj", false, true ' Tkdlqjrj is protection password
	Call keyset
End Sub

Private Sub Workbook_BeforeClose(Cancel)
	Call unkeyset
End Sub

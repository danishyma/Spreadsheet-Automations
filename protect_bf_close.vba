Private Sub Workbook_BeforeClose(Cancel As Boolean)

  ' protects the spreadsheet and workbook before closing
  Sheets("Sheet").Protect Password:="password"
  ActiveWorkbook.Protect Password:="passwords"
  
  ' if the file was being used by a specific user a backup is created at an specific folder on internal network
  If Environ("username") = "a_user" Then
    ActiveWorkbook.SaveCopyAs (Environ("UserProfile") & "\Backup\Spreadsheet Copy - Controle de Ticket " & Format(Now(), "yyyy-mm-dd hh.nn.ss") & ".xlsm")
  Else
  End If
  
  Worksheets("Sheet").Protect Password:="password", DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
    :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
    AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
    AllowUsingPivotTables:=True


  ' sets this specific sheet visibility to hidden
  Sheets("Sheet").Visible = xlSheetVeryHidden
  
  End Sub
  
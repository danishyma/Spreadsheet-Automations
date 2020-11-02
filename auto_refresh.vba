Sub Auto_Open()

' Refresh data from pivot and elsewhere on open
  ActiveWorkbook.RefreshAll

  ' Sets the sheet to hidden as default check the users and set visibility accordingly
  Sheets("Sheet One").Visible = xlSheetVeryHidden
  
    'if a_user creates a backup of the file on a backup folder from internal network
    If Environ("username") = "a_user" Then
      ActiveWorkbook.SaveCopyAs (Environ("UserProfile") & "\Backup\Spreadsheet Copy " & Format(Now(), "yyyy-mm-dd hh.nn.ss") & ".xlsm")
    Else
    End If

    'if owner or b_user the workbook and spreadsheet to unprotected
    If Environ("username") = "main_user" Or Environ("username") = "b_user" Then
      Sheets("Sheet One").Unprotect Password:="somepassword"
      ActiveWorkbook.Unprotect Password:="somepassword"
    End If

    'sets the spreadsheet to visible to both case above
    Sheets("Sheet One").Visible = xlSheetVisible
    Sheets("Sheet One").Select

    Else
    End If

  'if is neither of the users above it requests a password otherwise it only displays limited information
  TextBox1.Show
    Password = InputBox("Enter the password to access Updates", "Password", "****")

    If Password = "password" Then
      Sheets("Sheet One").Visible = xlSheetVisible
      Sheets("Sheet One").Select

    Else: Sheets("Sheet One").Visible = xlSheetVeryHidden
      MsgBox "Access denied :("

    End If

End Sub
Sub AcompanhamentodeAtividades()

  ' if following users open the workbook 
  If Environ("username") = "sheet_owner" Or Environ("username") = "a_user" Then
    Workbooks.Open Filename:="F:\Folder\Spreadsheet.xlsm"
    Workbooks.Open Filename:="C:\Users\username\Desktop\Spreadsheetname - Backup.xlsm"

  ' activates a spreadsheet and unprotects the spreadsheet and the workbook
  Worksheets("Sheet One").Activate
  Sheets("Sheet One").Unprotect Password:="passoword"
  ActiveWorkbook.Unprotect Password:="passoword"

  ' add a new tab to the workbook
  ActiveWorkbook.Sheets.Add.Name = "Sheet Name"

  ' popluates the spreadsheet labels
  Worksheets("Sheet Name").Range("A1") = "Label 1"
  Worksheets("Sheet Name").Range("D1") = "Label 2"
  Worksheets("Sheet Name").Range("G1") = "Label 3"
  Worksheets("Sheet Name").Range("J1") = "Label 4"
  Worksheets("Sheet Name").Range("A2, D2, G2, J2") = "Label 5"

  Worksheets("Sheet Name").Range("M1") = "Label 6"
  Worksheets("Sheet Name").Range("M2") = "Label 7"
  Worksheets("Sheet Name").Range("N1") = "Label 8"
  Worksheets("Sheet Name").Range("N2") = "Label 9"

  Worksheets("Sheet Name").Range("P1") = "Label 10"
  Worksheets("Sheet Name").Range("P2") = "Label 11"
  Worksheets("Sheet Name").Range("Q1") = "Label 12"
  Worksheets("Sheet Name").Range("Q2") = "Label 613"

  ' goes to spreadsheet and runs across it checking if tickets and populates the new sheet if the tickets are not responded and accordingly to the period they are unanswered 
  Worksheets("Sheet One").Activate
    Range("D4").Select

    While (ActiveCell.Value <> "")

      If ActiveCell.Value >= CDate(Format(Now - (Weekday(Now(), vbThursday) - 1) - 6, "dd/mm/yyyy")) And _
      ActiveCell.Value <= CDate(Format(Now - (Weekday(Now(), vbThursday) - 1), "dd/mm/yyyy")) Then
      
        If ActiveCell.Offset(0, 13) = "" And ActiveCell.Offset(0, 11) >= 2 Or _
        ActiveCell.Offset(0, 13) = 1 And ActiveCell.Offset(0, 11) >= 3 Or _
        ActiveCell.Offset(0, 13) = 6 And ActiveCell.Offset(0, 11) >= 4 Or _
        ActiveCell.Offset(0, 13) = 7 And ActiveCell.Offset(0, 11) >= 4 Then _
        ActiveCell.Offset(0, -3).Resize(1, 2).Copy (Sheets("Sheet Name").Range("A1").End(xlDown).Offset(1))

      End If
      
      If ActiveCell.Value >= CDate(Format(Now - (Weekday(Now(), vbFriday) - 1) - 19, "dd/mm/yyyy")) And _
      ActiveCell.Value <= CDate(Format(Now - (Weekday(Now(), vbFriday) - 1) - 13, "dd/mm/yyyy")) Then

        If ActiveCell.Offset(0, 12) > 14 Then _
        ActiveCell.Offset(0, -3).Resize(1, 2).Copy (Sheets("Sheet Name").Range("D1").End(xlDown).Offset(1))
      
      End If
              
      If ActiveCell.Offset(0, 6) = "PD" Or ActiveCell.Offset(0, 6) = "SR" Then _
        If ActiveCell.Offset(0, 8) > 14 Then _
        ActiveCell.Offset(0, -3).Resize(1, 2).Copy (Sheets("Sheet Name").Range("G1").End(xlDown).Offset(1))
          
      If ActiveCell.Offset(0, 6) = "PD" Then _
        If ActiveCell.Offset(0, 8) > 7 Then _
        ActiveCell.Offset(0, -3).Resize(1, 2).Copy (Sheets("Sheet Name").Range("J1").End(xlDown).Offset(1))
      
      ActiveCell.Offset(1).Select

    Wend
 
    Sheets("Sheet").Activate
    Range("A16").Select
       
      While (ActiveCell.Value <> "")
              
        If ActiveCell.End(xlToRight) = "1" Then
        ActiveCell.Copy (Sheets("Sheet Name").Range("M1").End(xlDown).Offset(1))
        Sheets("Sheet Name").Range("N1").Copy (Sheets("Sheet Name").Range("N1").End(xlDown).Offset(1))
        End If
        If ActiveCell.End(xlToRight) = "0,5" Then
        ActiveCell.Copy (Sheets("Sheet Name").Range("M1").End(xlDown).Offset(1))
        Sheets("Sheet Name").Range("N2").Copy (Sheets("Sheet Name").Range("N1").End(xlDown).Offset(1))
        End If
        
        ActiveCell.Offset(1).Select
          
      Wend
       
    Sheets("Sheet Two").Activate
    Range("A29").Select
       
      While (ActiveCell.Value <> "")
              
        If ActiveCell.End(xlToRight) = "1" Then
        ActiveCell.Copy (Sheets("Sheet Name").Range("P1").End(xlDown).Offset(1))
        Sheets("Sheet Name").Range("Q1").Copy (Sheets("Sheet Name").Range("Q1").End(xlDown).Offset(1))
        End If
        If ActiveCell.End(xlToRight) = "0,5" Then
        ActiveCell.Copy (Sheets("Sheet Name").Range("P1").End(xlDown).Offset(1))
        Sheets("Sheet Name").Range("Q2").Copy (Sheets("Sheet Name").Range("Q1").End(xlDown).Offset(1))
        End If
        
        ActiveCell.Offset(1).Select
          
      Wend
  End If

End Sub

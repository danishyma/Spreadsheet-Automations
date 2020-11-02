Sub Macro_Criarlink()

  'turns screen updating off to speed up your macro code
  Windows.Application.ScreenUpdating = False

  'if the cel contains () witch was used to denote a second entry on the spreasheet of the same ticket the program copies the ticket number before the parenthesis and creates a link to the system
  If InStr(1, ActiveCell.Value, "(" & ")", vbTextCompare) = 0 Then

    While (ActiveCell.Value <> "")

      Worksheets("Tickets Recebidos").Hyperlinks.Add _
        Anchor:=ActiveCell, _
        Address:="http://address?ticketId=" + Split(ActiveCell.Value, "(")(0), _
        ScreenTip:="Ticket " + ActiveCell.Value, _
        TextToDisplay:=ActiveCell.Value

      ActiveCell.Offset(1).Select

    Wend

  Else

    'otherwise it creates the link directly
    While (ActiveCell.Value <> "")

      Worksheets("Tickets Recebidos").Hyperlinks.Add _
        Anchor:=ActiveCell, _
        Address:="http://address?ticketId=" + ActiveCell.Value, _
        ScreenTip:="Ticket " + ActiveCell.Value, _
        TextToDisplay:=ActiveCell.Value

      ActiveCell.Offset(1).Select

    Wend

  End If

End Sub



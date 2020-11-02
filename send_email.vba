Sub enviar e-mail
   
Dim appOutlook As Object
Dim outMail As Object
Dim TicketColumn As Range, TicketRow As Range, R As Range, C As Range
Dim str As String
 
Set OutApp = CreateObject("Outlook.Application")
Set outMail = OutApp.CreateItem(0)
   
' if the spreadsheet contains one type of data it will create this label that will be assigned later to the e-mail
If Workbooks("Spreadsheet One - " & Format(Date, "yyyy-mm-dd") & "_Pendentes.xlsx").Sheets("Geral").Range("Q19") > 0 Then
  label_one = "<br><br> Label Message: <br>"
End If

' if the spreadsheet contains one type of data it will create this label and table that will be assigned later to the e-mail
If Workbooks("Spreadsheet One - " & Format(Date, "yyyy-mm-dd") & "_Pendentes.xlsx").Sheets("Geral").Range("Q19") > 0 Then
  
  label_two = Workbooks("Spreadsheet Two - " & Format(Date, "yyyy-mm-dd") & "_Pendentes.xlsx").Sheets("Lista").Activate
    Set TicketColumn = Range("A3", Range("A1").End(xlDown))
    ' Creates the table labels
    str = "<table>" & "<th style=background-color:#17365D> Label1 <th style=background-color:#17365D> Label2" & _
    "<th style=background-color:#17365D> Label3 <th style=background-color:#17365D> Label 4"
    
    For Each R In TicketColumn
    str = str & "<tr style=background-color:#E0E0E0>"
    Set TicketRow = Range(R, R.End(xlToRight))
    
    For Each C In TicketRow
    str = str & "<td align=center>" & C.Value & "</td>"
    Next C
    str = str & "</tr>"
    Next R
    str = str & "</table>"
    label_two = str
    
End If

' Cretaed updated graphs that will be added on the email, the graph on the folder gets substituted every time the macro runs
ActiveWorkbook.Sheets("Sheet one").ChartObjects("Graph 2").Activate
Set Grafico = ActiveWorkbook.Sheets("Shee One").ChartObjects("Graph 2").Chart
Grafico.Export Filename:="C:\temp\Graph2.jpg", filtername:="JPG"

ActiveWorkbook.Sheets("Sheet one").ChartObjects("Graph 6").Activate
Set Grafico = ActiveWorkbook.Sheets("Sheet one").ChartObjects("Graph 6").Chart
Grafico.Export Filename:="C:\temp\Graph6.jpg", filtername:="JPG"

' It creates an email on outlook, it doesn't send in case the user wants to add any information but it is possble to send it automatically
  With outMail
  
    .BodyFormat = olFormatHTML
    
    .To = "Adressed to"
    .CC = "Adresses in copy"
    .BCC = "Hidden adresses in copy"
    .Subject = "[TAG] E-mail Subject - " + Format(Date, "yyyy-mm-dd")
          
    .display
                
    Set olInsp = .GetInspector
    Set wdDoc = olInsp.WordEditor
    wdDoc.Range.InsertBefore
    
    ' This is the e-mail content, you can add the date, info from the sheets and the graphs previously created
    .htmlbody = "<head style=font-family:calibri;font-size:14.5'> <body>" & _
      "E-mail content <b>" & Format(Now - (Weekday(Now(), vbFriday) - 1), "dddd (dd/mm)") & _
      "E-mail content <b>" & + CStr(Sheets("Geral").Range("F19")) & _
      "E-mail content <b>" & _
      "<br> <img src='C:\temp\Grafico6.jpg'>" & label_one & label_two & _
      "<p> E-mail content <b> <br>" & "<br> <img src='C:\temp\Grafico2.jpg'>" & _
      "<p> E-mail content <b>" & "</head></body>" & .htmlbody
    
    ' Adds this spreadsheets previously created to the e-mail           
    .Attachments.Add Environ("UserProfile") & "\Desktop\Spreadsheet - " & Format(Date, "yyyy-mm-dd") & "_Geral.xlsx"
    .Attachments.Add Environ("UserProfile") & "\Desktop\Spreadsheet- " & Format(Date, "yyyy-mm-dd") & "_Pendentes.xlsx"
    
  End With

  ' Closes the secondary spreadsheets used
  Workbooks("Spreadsheet - " & Format(Date, "yyyy-mm-dd") & "_Geral.xlsx").Worksheets("Tickets Recebidos").Activate
  ThisWorkbook.Close
  Workbooks("Spreadsheet - " & Format(Date, "yyyy-mm-dd") & "_Pendentes.xlsx").Worksheets("Tickets Recebidos").Activate
  ThisWorkbook.Close

Else: MsgBox ("Access Denied")

End If

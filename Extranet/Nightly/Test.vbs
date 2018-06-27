lfilename="SummaryReport1.xls"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim act 
    Set act = fso.OpenTextFile(lfilename, 2, TRUE)


	act.WriteLine("<?xml version=""1.0""?>")

	act.WriteLine("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""")
	act.WriteLine("xmlns:o=""urn:schemas-microsoft-com:office:office""")
	act.WriteLine("xmlns:x=""urn:schemas-microsoft-com:office:excel""")
	act.WriteLine("xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""")
	act.WriteLine("xmlns:html=""http://www.w3.org/TR/REC-html40"">")
	act.WriteLine("<OfficeDocumentSettings xmlns=""urn:schemas-microsoft-com:office:office"">")
	act.WriteLine("<DownloadComponents/>")
	act.WriteLine("<LocationOfComponents HRef=""file:///\\""/>")
	act.WriteLine("</OfficeDocumentSettings>")
	act.WriteLine("<ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel"">")
	act.WriteLine("<WindowHeight>12525</WindowHeight>")
	act.WriteLine("<WindowWidth>15195</WindowWidth>")
	act.WriteLine("<WindowTopX>480</WindowTopX>")
	act.WriteLine("<WindowTopY>120</WindowTopY>")
	act.WriteLine("<ActiveSheet>2</ActiveSheet>")
	act.WriteLine("<ProtectStructure>False</ProtectStructure>")
	act.WriteLine("<ProtectWindows>False</ProtectWindows>")
	act.WriteLine("</ExcelWorkbook>")
	act.WriteLine("<Styles>")
	act.WriteLine("<Style ss:ID=""Default"" ss:Name=""Normal"">")
	act.WriteLine("<Alignment ss:Vertical=""Bottom""/>")
	act.WriteLine("<Borders/>")
	act.WriteLine("<Font/>")
	act.WriteLine("<Interior/>")
	act.WriteLine("<NumberFormat/>")
	act.WriteLine("<Protection/>")
	act.WriteLine("</Style>")
	act.WriteLine("</Styles>")
	act.WriteLine("<Worksheet ss:Name=""Sheet1"">")
	act.WriteLine("<Table>")
	act.WriteLine("<Row>")
   	act.WriteLine("<Cell><Data ss:Type=""String"">Name</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Title</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Company</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Address</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Telephone</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Email</Data></Cell>")
	act.WriteLine("<Cell><Data ss:Type=""String"">Fax</Data></Cell>")
	act.WriteLine("</Row>")
	act.WriteLine("</Table>")
	act.WriteLine("</Worksheet>")
	
act.WriteLine("<Worksheet ss:Name=""Sheet3"">")
act.WriteLine("<Table ss:ExpandedColumnCount=""1"" ss:ExpandedRowCount=""1"" x:FullColumns=""1""")
act.WriteLine("x:FullRows=""1"">")
act.WriteLine("<Row>")
act.WriteLine("<Cell><Data ss:Type=""Number"">3</Data></Cell>")
act.WriteLine("</Row>")
act.WriteLine("</Table>")
act.WriteLine("</Worksheet>")
act.WriteLine("</Workbook>")
	act.close
	
Set JMail = CreateObject("JMail.SMTPMail")
JMail.ReturnReceipt = false
JMail.Silent=true
JMail.ClearAttachments
JMail.ServerAddress = "mail.fluke.com"


' Send mail to Peter and Kelly

JMail.ClearRecipients
JMail.AddRecipient "jigar.joshi@fluke.com"
JMail.SenderName = "Summary Reports"
JMail.Sender = "WebMail@fluke.com"
JMail.ContentType = "text/plain"
JMail.ContentTransferEncoding ="quoted-printable"
JMail.AddNativeHeader "X-MimeOLE:Produced fluke.com"
JMail.ISOEncodeHeaders= true
JMail.Subject = "This is a Test of JMail"
JMail.AddAttachment "SummaryReport1.xls"
JMail.Send

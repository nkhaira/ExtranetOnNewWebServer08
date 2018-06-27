Imports System.Data.SqlClient
Imports System.Text
Imports System.collections
Imports System.data
Namespace ExcelExport

Partial Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub


    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sqlConnReports As New SqlConnection
        Dim sqlCmdReport As New SqlCommand
        'Dim Asset_ID_Old As Long
        Dim arrId As New ArrayList
        Dim objArrPrmLog1 As New SqlParameter("@siteid", SqlDbType.Int)
        Dim objArrPrmLog2 As New SqlParameter("@language", SqlDbType.VarChar, 100)
        Dim objArrPrmLog3 As New SqlParameter("@categoryid", SqlDbType.Int)
        Dim objArrPrmLog4 As New SqlParameter("@GroupId", SqlDbType.VarChar, 300)
        Dim objArrPrmLog5 As New SqlParameter("@Country", SqlDbType.VarChar, 100)
        Dim objArrPrmLog6 As New SqlParameter("@Submitted_By", SqlDbType.Int)
        Dim objArrPrmLog7 As New SqlParameter("@Campaign", SqlDbType.Int)
        Dim objArrPrmLog8 As New SqlParameter("@SortBy", SqlDbType.Int)

        Dim sqlHtml As New StringBuilder("")

        Dim strProcName As String = "getassets"
        Dim i As Int16
        Dim strServerName As String
        Try

            strServerName = UCase(Request.ServerVariables("SERVER_NAME"))

            If InStr(strServerName, "DTMEVTVSDV15") > 0 Then
                   sqlConnReports.ConnectionString = "data source=EVTIBG18.TC.FLUKE.COM;Initial Catalog=fluke_sitewide;persist security info=false;user id=FLUKEWEBUSER;Password=57OrcaSQL;"
            ElseIf InStr(strServerName, "DEV") > 0 Then
                   sqlConnReports.ConnectionString = "data source=EVTIBG18.TC.FLUKE.COM;Initial Catalog=fluke_sitewide;persist security info=false;user id=FLUKEWEBUSER;Password=57OrcaSQL;"
            ElseIf InStr(strServerName, "TEST") > 0 Then
                   sqlConnReports.ConnectionString = "data source=FLKTST18.DATA.ib.fluke.com;Initial Catalog=fluke_sitewide;persist security info=false;user id=FLUKEWEBUSER;Password=57OrcaSQL;"
            ElseIf InStr(strServerName, "PRD") > 0 Then
                    sqlConnReports.ConnectionString = "data source=FLKPRD18.DATA.ib.fluke.com;Initial Catalog=fluke_sitewide;persist security info=false;user id=FLUKEWEBUSER;Password=57OrcaSQL;"
            Else
              'Default to PRODUCTION
                 sqlConnReports.ConnectionString = "data source=FLKPRD18.DATA.ib.fluke.com;Initial Catalog=fluke_sitewide;persist security info=false;user id=FLUKEWEBUSER;Password=57OrcaSQL;"
            End If

            sqlConnReports.Open()
            sqlCmdReport.Connection = sqlConnReports
            objArrPrmLog1.Value = Request.QueryString("Site_ID")
            If Request.QueryString("language") <> "" Then
                objArrPrmLog2.Value = Request.QueryString("language")
            End If

            If Request.QueryString("categoryid") > 0 Then
                objArrPrmLog3.Value = Request.QueryString("categoryid")
            End If

             If Request.QueryString("GroupId") <> "" Then
                objArrPrmLog4.Value = Request.QueryString("GroupId")
            End If

            If Request.QueryString("Country") <> "" Then
                objArrPrmLog5.Value = Request.QueryString("Country")
            End If

            If Request.QueryString("Submitted_By") > 0 Then
                objArrPrmLog6.Value = Request.QueryString("Submitted_By")
            End If

            If Request.QueryString("Campaign") > 0 Then
                objArrPrmLog7.Value = Request.QueryString("Campaign")
            End If

            If Request.QueryString("SortBy") > 0 Then
                objArrPrmLog8.Value = Request.QueryString("SortBy")
            End If

            sqlCmdReport.Parameters.Add(objArrPrmLog1)
            sqlCmdReport.Parameters.Add(objArrPrmLog2)
            sqlCmdReport.Parameters.Add(objArrPrmLog3)
            sqlCmdReport.Parameters.Add(objArrPrmLog4)
            sqlCmdReport.Parameters.Add(objArrPrmLog5)
            sqlCmdReport.Parameters.Add(objArrPrmLog6)
            sqlCmdReport.Parameters.Add(objArrPrmLog7)
            sqlCmdReport.Parameters.Add(objArrPrmLog8)
            'Debugging Code
            'Response.Write(objArrPrmLog1.Value & "<br>")
            'Response.Write(objArrPrmLog2.Value & "<br>")
            'Response.Write(objArrPrmLog3.Value & "<br>")
            'Response.Write(objArrPrmLog4.Value & "<br>")
            'Response.Write(objArrPrmLog5.Value & "<br>")
            'Response.Write(objArrPrmLog6.Value & "<br>")
            'Response.Write(objArrPrmLog7.Value & "<br>")
            'Response.Write(objArrPrmLog8.Value & "<br>")
            'Response.End()
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            sqlCmdReport.CommandType = CommandType.StoredProcedure
            sqlCmdReport.CommandText = "getassets"
            Dim objReader As SqlDataReader
            sqlHtml.Append("<table border=1>")
            Dim iCol As Integer

            sqlHtml.Append("<tr>")
            objReader = sqlCmdReport.ExecuteReader(CommandBehavior.CloseConnection)
            sqlHtml.Append("<tr>")
            If Request.QueryString("Site_ID") = 82 Then
                For iCol = 0 To objReader.FieldCount - 1
                    sqlHtml.Append("<td ><b>" & UnicodeHtmlEncode(objReader.GetName(iCol).ToString()) & "</b></td>")
                Next
            Else
                For iCol = 0 To 25
                    sqlHtml.Append("<td ><b>" & UnicodeHtmlEncode(objReader.GetName(iCol).ToString()) & "</b></td>")
                Next
            End If
            sqlHtml.Append("</tr>")

            If objReader.HasRows = True Then
                While (objReader.Read)
                    If arrId.IndexOf(objReader.Item("id")) < 0 Then
                       sqlHtml.Append("<tr>")
                       If Request.QueryString("Site_ID") = 82 Then
                           For i = 0 To objReader.FieldCount - 1
                               If IsDBNull(objReader.Item(i)) = False Then
                                   sqlHtml.Append("<td>" & UnicodeHtmlEncode(Server.HtmlEncode(objReader.Item(i))) & "</td>")
                               Else
                                   sqlHtml.Append("<td>" & "" & "</td>")
                               End If
                           Next
                       Else
                           For i = 0 To 25
                               If IsDBNull(objReader.Item(i)) = False Then
                                   sqlHtml.Append("<td>" & UnicodeHtmlEncode(Server.HtmlEncode(objReader.Item(i))) & "</td>")
                               Else
                                   sqlHtml.Append("<td>" & "" & "</td>")
                               End If
                           Next
                       End If
                       sqlHtml.Append("</tr>")
                       'Asset_ID_Old = objReader.Item("id")
                       arrId.Add(objReader.Item("id"))
                     End If
                End While
            End If
            sqlHtml.Append("</table>")
            Session("strHTMLReport") = sqlHtml.ToString()
            Response.Redirect("Excelreport.aspx")
        Catch objEx As Exception
            Throw objEx
        Finally
            sqlConnReports.Dispose()
            sqlCmdReport.Dispose()
        End Try
    End Sub
    Public Function UnicodeHtmlEncode(ByVal text As String) As String
        Dim chars As Char() = HttpUtility.HtmlEncode(text).ToCharArray()
        Dim result As New StringBuilder(text.Length + CInt((text.Length * 0.1)))

        For Each c As Char In chars
            Dim value As Integer = Convert.ToInt32(c)
            If value > 127 Then
                result.AppendFormat("&#{0};", value)
            Else
                result.Append(c)
            End If
        Next

        Return result.ToString()
    End Function
End Class

End Namespace

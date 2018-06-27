
Namespace ExcelExport

Partial Class ExcelReport
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

        'This will make the browser interpret the output as an Excel file
        Response.AddHeader("Content-Disposition", "filename=Data.xls")
        ' Set the content type to Excel.
        Response.ContentType = "application/vnd.ms-excel"

        ' Remove the charset from the Content-Type header.
        Response.Charset = "UTF-16"

        ' Turn off the view state.
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter

        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        Dim ltrCtrl As New Literal

        ltrCtrl.Text = Session("strHTMLReport")
        ltrCtrl.RenderControl(hw)

        ' Write the HTML back to the browser.
        Response.Write(tw.ToString())

        ' End the response.
        Response.End()

    End Sub
#Region "User defined functions and procedures"

#End Region
End Class

End Namespace

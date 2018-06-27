using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;


namespace ExtranetSW_Stream_IT
{
	/// <summary>
	/// Summary description for FindIt.
	/// </summary>
	public partial class SW_Stream_IT : System.Web.UI.Page
	{
		protected void Page_Load(object sender, System.EventArgs e)
		{
			string strVirtualFilepath = Request["filepath"];
            string strContentType = Request["contenttype"].Trim();
            
            System.IO.Stream iStream = null;

            // Buffer to read 10K bytes in chunk:
            byte[] buffer = new Byte[10000];

            // Length of the file:
            int length;

            // Total bytes to read:
            long dataToRead;

            // Identify the file to download including its path.
            string filepath = Server.MapPath(strVirtualFilepath);

            // Identify the file name.
            string filename = System.IO.Path.GetFileName(filepath);

            if (strVirtualFilepath != "")
            {
                try
                {
                    // Open the file.
                    iStream = new System.IO.FileStream(filepath, System.IO.FileMode.Open,
                                System.IO.FileAccess.Read, System.IO.FileShare.Read);


                    // Total bytes to read:
                    dataToRead = iStream.Length;

                    Response.AddHeader ("Content-Length", dataToRead.ToString());
                    
                    if (strContentType != "")
                    {
                        Response.ContentType = strContentType;
                        switch (strContentType)
                        {
                            case "application/octet-stream":                                
                                Response.AddHeader("Content-Disposition", "attachment; filename=" + filename);  //attachment
                                break;
                            default:
                                Response.AddHeader("Content-Disposition", "inline; filename=" + filename);
                                break;
                        }
                    }
                    else
                    {

                        Response.ContentType = "application/octet-stream";
                        Response.AddHeader("Content-Disposition", "attachment; filename=" + filename);  //attachment
                    
                    }

                    // Read the bytes.
                    while (dataToRead > 0)
                    {
                        // Verify that the client is connected.
                        //if (Response.IsClientConnected)
                        //{
                            // Read the data in buffer.
                            length = iStream.Read(buffer, 0, 10000);

                            // Write the data to the current output stream.
                            Response.OutputStream.Write(buffer, 0, length);

                            // Flush the data to the HTML output.
                            Response.Flush();

                            buffer = new Byte[10000];
                            dataToRead = dataToRead - length;
                        //}
                        //else
                        //{
                            //prevent infinite loop if user disconnects
                        //    dataToRead = -1;
                        //}
                    }
                }
                catch (Exception ex)
                {
                    // Trap the error, if any.
                    Response.Write("Error : " + ex.Message);
                }
                finally
                {
                    if (iStream != null)
                    {
                        //Close the file.
                        iStream.Close();                        
                    }
                }
            }

		}
		

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
		}
		#endregion
	}
}

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
using DanaherTM.ProductEngine;
using DanaherTM.Framework.ExceptionHandling;

namespace ExtranetPcat
{
	/// <summary>
	/// Summary description for test.
	/// </summary>
	public class test : System.Web.UI.Page
	{
        protected System.Web.UI.WebControls.DataGrid dbAssets;
        private DanaherTM.ProductEngine.ProductEngineInstance objPEE;
        private DanaherTM.ProductEngine.Catalog objCatalog;
        protected string LinkTitle;
    
		private void Page_Load(object sender, System.EventArgs e)
		{
            string strPID = "50004";
            Product objProd;
            try
            {
                objPEE = new ProductEngineInstance("en-us",DateTime.Now);
                objCatalog = objPEE.Catalogs["Fnet-US"];
                
                objProd =   objCatalog.Products[strPID];            
                //Set the category code for Product Family control
                CategoryLocalized objCatLocalized = objProd.DefaultCategory.Localized;
                                    
                //Set grid headers
                LinkTitle = "Download";
                dbAssets.Columns[0].HeaderText= "Title";
                dbAssets.Columns[2].HeaderText= "DOWNLOAD OPTIONS";

                //Assign datasource for the grid
                dbAssets.DataSource = getAssets(objProd);
                dbAssets.DataBind();
                
            }
            catch (ProductEngineException ex)
            {
                if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.CommonControls, ex))
                {   
                    throw;
                }
            }
            catch (NullReferenceException ex)
            {
                if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.CommonControls, ex))
                {   
                    throw;
                }
            }
            catch (Exception ex)
            {
                if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.CommonControls, ex))
                {   
                    throw;
                }
            }
            finally
            {                
                objPEE.Dispose();                
            }                    
		}


        #region User defined functions

        private DataView getAssets(DanaherTM.ProductEngine.Product objProd)
        {            
            DanaherTM.ProductEngine.ProductLocalized objProdLocalized;
            DataTable dtAssets = new DataTable();
            dtAssets.Columns.Add("AssetName");
            dtAssets.Columns.Add("Description");
            dtAssets.Columns.Add("PID");
            dtAssets.Columns.Add("OracleID"); 
            dtAssets.Columns.Add("FileSize");
            try
            {
                DanaherTM.ProductEngine.Products objRelatedProds = objProd.AssetsBySubType(DBUtils.DataStructures.ProductSubTypes.Manuals);
                foreach (DanaherTM.ProductEngine.Product objAsset in objRelatedProds)
                {                
                    objProdLocalized = objAsset.Localized;
                    dtAssets.Rows.Add(CreateNewRow(objAsset.Localized.Name,
                        objAsset.Localized.LongDescription,
                        objAsset.ID.ToString(),
                        objAsset.Localized.OraclePartNum, 
                        objAsset.Localized.FileSize,
                        dtAssets));            
                }
            }
            catch (ProductEngineException ex)
            {
                ExceptionPolicy.HandleException(ExceptionInstance.FlukeNetworks.ProductEngine,ex); 
            }
            catch (DataException ex)
            {
                ExceptionPolicy.HandleException(ExceptionInstance.FlukeNetworks.ProductEngine,ex);
            }
            catch (Exception ex)
            {
                ExceptionPolicy.HandleException(ExceptionInstance.FlukeNetworks.WebPages,ex);                  
            }
            
            DataView dv = dtAssets.DefaultView;
            dv.Sort="AssetName";
            return dv;            
        }

     
        private DataRow CreateNewRow(string shortdesc,string longdesc,string pid,string oranum, string FileSize, DataTable dt)
        {            
            DataRow dr = dt.NewRow();

            dr[0] = shortdesc;
            dr[1] = longdesc;
            dr[2] = pid;
            dr[3] = oranum;
            dr[4] = FileSize;

            return dr;
        }
        #endregion
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
            this.Load += new System.EventHandler(this.Page_Load);

        }
		#endregion
	}
}

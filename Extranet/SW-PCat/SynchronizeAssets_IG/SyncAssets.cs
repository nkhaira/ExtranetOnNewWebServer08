using System;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Scripting;
using DanaherTM.Framework.ExceptionHandling;
using System.Web;
using System.Text;
using System.Data.Sql;
using System.Data.SqlTypes;
using Framework.Modules.PCAT.API;
using System.Net.Mail;

namespace SynchronizeAssets
{
	public class SyncAssets
	{
		public SyncAssets()
		{
			///Constructor;
		}
		/*
        //public ArrayList GetLocales(string strLanguage)
        //{
        //    Locales LanguageLocales=new  Locales(strLanguage);
        //    ArrayList LocaleArray=new ArrayList();
        //    foreach(Locale LangLocale in LanguageLocales)
        //    {
        //        LocaleArray.Add(LangLocale.LocaleValue);
        //    }
        //    return LocaleArray;
        //}
		*/
		
		public void updateLocalizedtable()
		{	
				SqlDataReader assetDataReader=null;
				SqlConnection oSiteWideDB=null;
				StringBuilder strResult=new StringBuilder("");
				string smtpServer="";
				string mailRecepients="";
				string fromUser="";
				string strProductId="";
                string strConnSiteWide = "";
                string strConnPCat = "";
			    bool boolMailYn = false;

				smtpServer=System.Configuration.ConfigurationSettings.AppSettings.Get("SmtpServer");
				mailRecepients=System.Configuration.ConfigurationSettings.AppSettings.Get("Recepients");
				fromUser=System.Configuration.ConfigurationSettings.AppSettings.Get("From");
                ////boolMailYn=Convert.ToBoolean(System.Configuration.ConfigurationSettings.AppSettings.Get("MailYN"));
                strConnSiteWide = System.Configuration.ConfigurationSettings.AppSettings.Get("SiteWide");
                strConnPCat = System.Configuration.ConfigurationSettings.AppSettings.Get("ConnStr_PCAT");
                DataTable dtProducts = new DataTable();

				/*  //Get the assets from sitewide DB */
				try
				{
                    int iProductId = 0;
					//Create database object using enterprise library i.e(Dataconfig.config)
                    oSiteWideDB = new SqlConnection(strConnSiteWide);
					//Get the data reader for assets
				    SqlCommand cmdToExecute = new SqlCommand();
                    cmdToExecute.Connection = oSiteWideDB;
                    string sSQL = "";
                    sSQL = "PCAT_FLUKE_SYNCASSETS";
                    try
                    {
                        cmdToExecute.CommandText = sSQL;
                        cmdToExecute.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataAdapter adapter = new SqlDataAdapter(cmdToExecute);
                        cmdToExecute.Parameters.Clear();
                        adapter.Fill(dtProducts);
                    }
                    catch (Exception ex)
                    {}
                    finally
                    {}

                    ////assetDataReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.StoredProcedure,"PCAT_FLUKE_SYNCASSETS");
                    DataTable dtLngLoc = GetLanguageLocale();
					foreach(DataRow drProduct in dtProducts.Rows)
					{
                        DataRow[] drLocls = dtLngLoc.Select("ISO2='" + drProduct["Language"].ToString().Substring(0, 2) + "'");
						////localArray=GetLocales(assetDataReader.GetValue(0).ToString().Substring(0,2));
						try
						{
                            iProductId = Convert.ToInt32(drProduct[2].ToString());
                            if (iProductId > 0)
                            {
                                Product ModifiedProduct = new Product(iProductId, "PCATAuthoring");
                                strProductId = ModifiedProduct.ID.ToString();
                                foreach (DataRow drLoc in drLocls)
                                {
                                    try
                                    {
                                        ProductLocalized ModifiedLocalProduct = new ProductLocalized(iProductId, drLoc["Locale"].ToString(), "PCATAuthoring");
                                        if (drProduct[3].ToString() == "1")
                                        {
                                            ModifiedLocalProduct.StartDate = Convert.ToDateTime(drProduct[4].ToString());
                                        }
                                        else
                                        {
                                            ModifiedLocalProduct.StartDate = new DateTime(2050, 1, 1);
                                        }

                                        ModifiedLocalProduct.Save();
                                        strResult.Append("Asset Item Number=" + ModifiedLocalProduct.OraclePartNum + ":StartDate=" + ModifiedLocalProduct.StartDate + "\n");
                                    }
                                    catch (Exception ex)
                                    {
                                        strResult.Append("Product Id=" + ModifiedProduct.ID + ":Error occurred" + "\n");
                                    }
                                }
                            }
						}
						catch(Exception ex)
						{
								strResult.Append("Product Id=" + strProductId + ":Error occurred"+ "\n");
						}
					}

					assetDataReader=null;
					oSiteWideDB=null;
				
					if (boolMailYn  == true)
					{
						SendEmail(mailRecepients,"Synchronization:Process completed successfully-" + DateTime.Now.ToString(),strResult.ToString() ,smtpServer,fromUser,"");   
					}
					
				}
				catch(Exception ex)
				{
					try
					{
						if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
						{   
							SendEmail(mailRecepients,"Synchronization:Error occurred while processing.-" + ex.Message + DateTime.Now.ToString(),strResult.ToString() ,smtpServer,fromUser,"");   
						}
					}
					catch(Exception LoggingException)
					{
						SendEmail(mailRecepients,"Synchronization:Error occurred while processing.-" + LoggingException.Message + DateTime.Now.ToString(),strResult.ToString() ,smtpServer,fromUser,"");   						
					}
				}
		}
		private void SendEmail(string strRecipient, string strSubject, string strBody, string strServer, 
                               string strSender,string Bugstate)
		{
			try
			{
               MailMessage myMessage = new MailMessage();

                myMessage.To.Add(new MailAddress(strRecipient));
                myMessage.From = new MailAddress(strSender);
                myMessage.Priority = MailPriority.Normal;
                myMessage.Subject = strSubject;
                myMessage.Body = strBody;

                try
                {
                    //send the message
                    SmtpClient smtp = new SmtpClient(strServer);
                    smtp.Send(myMessage);
                }
                catch (System.Exception ex)
                {
                } 
			
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}  
		}

        private DataTable GetLanguageLocale()
        {
            SqlConnection oMainConnection = new SqlConnection(System.Configuration.ConfigurationSettings.AppSettings["ConnStr_PCAT"].ToString().Trim());
            SqlCommand cmdToExecute = new SqlCommand();
            cmdToExecute.Connection = oMainConnection;
            string sSQL = "";
            
            sSQL = "SELECT LocaleID, ISO2, Locale, ISO3 FROM Locales WHERE (BrandID = 2)";
            DataTable oToReturn = new DataTable();
            try
            {
                cmdToExecute.CommandText = sSQL;
                cmdToExecute.CommandType = System.Data.CommandType.Text;
                SqlDataAdapter adapter = new SqlDataAdapter(cmdToExecute);
                cmdToExecute.Parameters.Clear();

                oMainConnection.Open();
                adapter.Fill(oToReturn);
               
            }
            catch (Exception ex)
            {
                //LogEntry.WriteLogEntry(ex.Message);
            }
            finally
            {
                if (oMainConnection.State != ConnectionState.Closed)
                {
                    oMainConnection.Close();
                }
            }

            return oToReturn;
        }
	}
}

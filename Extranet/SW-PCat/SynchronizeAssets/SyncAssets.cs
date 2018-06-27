using System;
using DanaherTM.ProductEngine;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Scripting;
using DataLibrary = Microsoft.Practices.EnterpriseLibrary.Data;
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;  
using Microsoft.Practices.EnterpriseLibrary.Configuration;
using DanaherTM.Framework.ExceptionHandling;
using System.Web.Mail;
using System.Text;
namespace SynchronizeAssets
{
	public class SyncAssets
	{
		public SyncAssets()
		{
			//Constructor;
		}
		
		public ArrayList GetLocales(string strLanguage)
		{
			Locales LanguageLocales=new  Locales(strLanguage);
			ArrayList LocaleArray=new ArrayList();
			foreach(Locale LangLocale in LanguageLocales)
			{
				LocaleArray.Add(LangLocale.LocaleValue);
			}
			return LocaleArray;
		}
				
		public void updateLocalizedtable()
		{	
			
				SqlDataReader assetDataReader=null;
				DataLibrary.Database oSiteWideDB=null;
				StringBuilder strResult=new StringBuilder("");
				string smtpServer="";
				string mailRecepients="";
				string fromUser="";
				string strProductId="";
			 bool boolMailYn=true;

				smtpServer=System.Configuration.ConfigurationSettings.AppSettings.Get("SmtpServer");
				mailRecepients=System.Configuration.ConfigurationSettings.AppSettings.Get("Recepients");
				fromUser=System.Configuration.ConfigurationSettings.AppSettings.Get("From");
				boolMailYn=Convert.ToBoolean(System.Configuration.ConfigurationSettings.AppSettings.Get("MailYN"));

				//Get the assets from sitewide DB
				try
				{
					//Create database object using enterprise library i.e(Dataconfig.config)
					oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
					//Get the data reader for assets
					assetDataReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.StoredProcedure,"PCAT_FNET_SYNCASSETS");     
			
					//Update the products in product engine
					ArrayList localArray;
					while(assetDataReader.Read())
					{
						localArray=GetLocales(assetDataReader.GetValue(0).ToString().Substring(0,2));
						try
						{
						Product ModifiedProduct=new Product(Convert.ToInt32(assetDataReader.GetValue(2).ToString()));
						strProductId =ModifiedProduct.ID.ToString();
						foreach(string LangLocale in localArray)
							{   
								try
								{
											ProductLocalized ModifiedLocalProduct = new ProductLocalized(ModifiedProduct,LangLocale);
											if (assetDataReader.GetValue(3).ToString()=="1")
											{
												  ModifiedLocalProduct.StartDate=Convert.ToDateTime(assetDataReader.GetValue(4).ToString());
											}
											else
											{
												  ModifiedLocalProduct.StartDate=DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never;
											}
									  ModifiedLocalProduct.Save();
											strResult.Append("Asset Item Number=" + ModifiedLocalProduct.OraclePartNum + ":StartDate=" + ModifiedLocalProduct.StartDate + "\n");
										}
										catch(ProductEngineException ex)
										{
											if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
											{   
												strResult.Append("Product Id=" + ModifiedProduct.ID + ":Error occurred" + "\n");
											}
										}
										catch(Exception ex)
										{
											if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
											{   
												strResult.Append("Product Id=" + ModifiedProduct.ID + ":Error occurred"+ "\n");
											}
								}
							}
						}
						catch(ProductEngineException ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
							{   
								strResult.Append("Product Id=" + strProductId + ":Error occurred"+ "\n");
							}
						}
						catch(Exception ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
								strResult.Append("Product Id=" + strProductId + ":Error occurred"+ "\n");
							}
						}
					}

					//Dispose the unused objects
					assetDataReader.Close();
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
		private void SendEmail(string strRecipient, string strSubject, string strBody, string strServer, string strSender,string Bugstate)
		{
			try
			{
				SmtpMail.SmtpServer =strServer;
				MailMessage mail = new MailMessage();	
				mail.To = strRecipient;
				mail.From = strSender;
				mail.Subject = strSubject;
				mail.BodyFormat=MailFormat.Text;   
				mail.Body = mail.Body + strBody; 
				
				if(strRecipient=="" || strSender=="") 
				{
					return;
				}
				else
				{
					SmtpMail.Send(mail);
				}
				
			
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}  
		}
	}
}

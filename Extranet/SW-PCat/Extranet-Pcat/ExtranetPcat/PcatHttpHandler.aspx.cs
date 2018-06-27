using System;
using System.Collections;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.IO;
using System.Text;
using DanaherTM.Framework.ExceptionHandling;
using System.Data.SqlClient;
using DataLibrary = Microsoft.Practices.EnterpriseLibrary.Data;
using System.Xml;
using DanaherTM.ProductEngine;
using System.Data;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Configuration;
namespace ExtranetPcat
{   
	public class PcatInterface : System.Web.UI.Page
	{
		private int Site_Id;
		//This page gets called from different Extranet asp  pages through XMLHttp.
		//Based on the operation flag passed corresponding routine gets executed.
		private void Page_Load(object sender, System.EventArgs e)
		{ 
			Site_Id=Convert.ToInt16(System.Configuration.ConfigurationSettings.AppSettings["Site_Id"]);

			if (Request.Form["operation"]=="P") 
			{
						RetrieveProducts();
			}
			else if(Request.Form["operation"]=="A" || Request.Form["operation"]=="U") 
			{
						AddUpdateAsset();
			}
			else if(Request.Form["operation"]=="D")
			{
						DeleteAsset();
			}
			else if(Request.Form["operation"]=="V")
			{
					 ValidateData();
			}
			else if(Request.Form["operation"]=="PA")
			{
						RetrieveProductsNotLinked();
			}
			else if(Request.Form["operation"]=="PAssets")
			{
						RetrieveProductsAssets();		
			}
			else if(Request.Form["operation"]=="Associate")
			{
					 AssociateAssets();	
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
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
		public void RetrieveProducts()
			//**********RetrieveProducts**********
			//NAME           : RetrieveProducts
			//PURPOSE        : This function is used to retrive products.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
							Response.ContentType = "text/xml";
							Response.Charset     = "utf-16";
							
							XmlHttpHandler HandleXmlRequest=new XmlHttpHandler();
							try
							{
										string Products;
										if (Request.Form["assetpid"]!="")
										{
											Products=HandleXmlRequest.EnemerateProducts(Convert.ToInt32(Request.Form["assetpid"]));
										}
										else
										{
											Products=HandleXmlRequest.EnemerateProducts(0);
										}
										Response.BinaryWrite(StringToBytes(Products));
										HandleXmlRequest=null;
							}
							catch(Exception ex)
									{
												if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
												{  
													//throw;
												}
									}
		}
		public void AddUpdateAsset()
			//**********AddUpdateAsset**********
			//NAME           : AddUpdateAsset
			//PURPOSE        : This function is used to Add or update asset info.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
			try
			{
							XmlHttpHandler AssetDocs=new XmlHttpHandler();
							string strproducts;
							string splitcharater;
							splitcharater=",";
							string[] ProductArray;
							string PID;
							//string calendarSql;
							string strIncludeExclude;

							DataLibrary.Database oSiteWideDB=null;
																	
							SqlDataReader assetDataReader=null;

							strproducts=Request.Form["products"];
							ProductArray=strproducts.Split(splitcharater.ToCharArray()[0]);
								
							PID=AssetDocs.CreateModifyAsset(Convert.ToBoolean(Request.Form["isclone"]),Request.Form["assetpid"],
								Request.Form["title"],Request.Form["description"],Request.Form["Filename"],Request.Form["FileSize"],
								Convert.ToDateTime(Request.Form["begindate"]),ProductArray,Request.Form["language"],
								Request.Form["operation"],Request.Form["Category_Type"],Request.Form["oraclenumber"],
								Request.Form["access"],Request.Form["industry"],Request.Form["IncludeExclude"],
								Convert.ToBoolean(Request.Form["status"]),Request.Form["oldLanguage"],Request.Form["oldItemNumber"],
								Convert.ToInt64(Request.Form["AssetId"]));
							try
							{
								oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
								assetDataReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text,"PCAT_FNET_ASSETCLONES_SEL " + Site_Id + "," + Request.Form["calendarId"].ToString());     
							}
							catch(Exception connectionEx)
							{
								if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, connectionEx))
								{   
									//throw;
								}		
							}
							if (assetDataReader != null)
							{
								while(assetDataReader.Read())
								{
									if (assetDataReader.GetValue(16).ToString().StartsWith("0")== true)
									{
										strIncludeExclude=assetDataReader.GetValue(16).ToString();
									}
									else if (assetDataReader.GetValue(16).ToString()=="none")
									{
										strIncludeExclude="none";
									}
									else 
									{
										strIncludeExclude= "1" + assetDataReader.GetValue(16).ToString();
									}
									XmlDocument xmlDomDocument =new XmlDocument();
									string productId;
									xmlDomDocument.LoadXml(PID);
									productId=xmlDomDocument.DocumentElement.InnerText.ToString();
									xmlDomDocument=null;

									string strFilePathName = "";
										
									if (Convert.ToInt32(assetDataReader.GetValue(17)) <=0)
									{
										strFilePathName = assetDataReader.GetValue(6).ToString();
										if (strFilePathName.LastIndexOf(@"\")>1)
										{
											strFilePathName = strFilePathName.Substring( strFilePathName.LastIndexOf(@"\")+ 1);
										}
										if (strFilePathName.LastIndexOf("/")>1)
										{
											strFilePathName = strFilePathName.Substring( strFilePathName.LastIndexOf("/")+ 1);
										}
													
										PID=AssetDocs.CreateModifyAsset(Convert.ToBoolean(assetDataReader.GetValue(11)),
											productId,assetDataReader.GetValue(0).ToString(),
											assetDataReader.GetValue(4).ToString(), strFilePathName,
											assetDataReader.GetValue(7).ToString(),
											Convert.ToDateTime(assetDataReader.GetValue(8).ToString()),ProductArray,
											assetDataReader.GetValue(5).ToString(),"U",assetDataReader.GetValue(1).ToString(),
											assetDataReader.GetValue(14).ToString(),assetDataReader.GetValue(2).ToString(),
											Request.Form["industry"],strIncludeExclude,Convert.ToBoolean(Request.Form["status"])
											,assetDataReader.GetValue(5).ToString(),assetDataReader.GetValue(14).ToString()
											,Convert.ToInt64(assetDataReader.GetValue(15)));
									}
									strIncludeExclude="";
								}
							}

							Response.ContentType = "text/xml";
							Response.Charset     = "utf-16";
							Response.BinaryWrite(StringToBytes(PID));
							AssetDocs=null;
						}
						catch(Exception ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
								//throw;
							}
						}
		}
		public void DeleteAsset()
			//**********DeleteAsset**********
			//NAME           : DeleteAsset
			//PURPOSE        : This function is used to delete asset info.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
						IDbTransaction objTrans = null;
						Database objData = DBUtils.GetDB();
						IDbConnection objIDbConn = objData.GetConnection();
						objIDbConn.Open();
						objTrans = objIDbConn.BeginTransaction();
						try
						{
								
							XmlHttpHandler AssetDocs=new XmlHttpHandler();
							AssetDocs.DeleteAsset( Request.Form["assetpid"],Request.Form["language"],Request.Form["operation"],
							Convert.ToBoolean(Request.Form["isclone"]),Convert.ToBoolean(Request.Form["DeleteAll"]),
							Convert.ToBoolean(Request.Form["setRelationship"]),
							Request.Form["itemNumber"],objTrans,null);

							if (objTrans != null)
							{
									objTrans.Commit();
									objTrans.Dispose();
									objIDbConn.Close();
									objIDbConn.Dispose();
									objData = null;
							}
						}
						catch(Exception ex)
						{
							if (objTrans != null)
							{
									objTrans.Rollback();
									objTrans.Dispose();
									objIDbConn.Close();
									objIDbConn.Dispose();
									objData = null;
							}

							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{
								//throw;
							}
						}
		}
		public void ValidateData()
			//**********ValidateData**********
			//NAME           : ValidateData
			//PURPOSE        : This function is used to validate asset info.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
						try
						{
							string functionResult;
							XmlHttpHandler validateLocale=new XmlHttpHandler();	
							functionResult=validateLocale.ValidateCatalogsLocales(Request.Form["IncludeExclude"],
							Request.Form["language"],Request.Form["prodSubType"],
							Request.Form["assetpid"],Request.Form["oraclenumber"]);
							Response.ContentType = "text/xml";
							Response.Charset     = "utf-16";
							Response.BinaryWrite(StringToBytes(functionResult));
							validateLocale=null;
						}
						catch(Exception ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
								//throw;
							}
						}
		}
		public void RetrieveProductsNotLinked()
			//**********RetrieveProductsNotLinked**********
			//NAME           : RetrieveProductsNotLinked
			//PURPOSE        : This function is used to retrieve products which are not linked to any assets.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
							try
							{
										string functionResult;
										XmlHttpHandler productAssets=new XmlHttpHandler();	
										functionResult = productAssets.GetProducts();
										Response.ContentType = "text/xml";
										Response.Charset     = "utf-16";
										Response.BinaryWrite(StringToBytes(functionResult));
										productAssets=null;
							}
							catch(Exception ex)
							{
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
										{   
											//throw;
										}
							}
		}
		public void RetrieveProductsAssets()
			//**********RetrieveProductsAssets**********
			//NAME           : RetrieveProductsAssets
			//PURPOSE        : This function is used to retrieve assets related to a specific product.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 16-04-2007 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
					try
					{
										string functionResult;
										XmlHttpHandler productAssets=new XmlHttpHandler();	
										functionResult = productAssets.GetProductAssets(Request.Form["assetPid"]);
										Response.ContentType = "text/xml";
										Response.Charset     = "utf-16";
										Response.BinaryWrite(StringToBytes(functionResult));
										productAssets=null;
					}
					catch(Exception ex)
					{
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
										{   
											//throw;
										}
					}	
		}
		public void AssociateAssets()
		//**********AssociateAssets**********
		//NAME           : AssociateAssets
		//PURPOSE        : This function is used to associate assets to a specific product.
		//PARAMETERS     : 
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 16-04-2007 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
			{
					try
					{
								string functionResult;
								XmlHttpHandler Associations=new XmlHttpHandler();	
								functionResult = Associations.SetRelationships(Request.Form["assetPid"],Request.Form["assetId"]);
								Response.ContentType = "text/xml";
								Response.Charset     = "utf-16";
								Response.BinaryWrite(StringToBytes(functionResult));
								Associations=null;
					}
					catch(Exception ex)
					{
								if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
								{   
									//throw;
								}
					}
			}
		public byte[] StringToBytes(string str)
			//**********StringToBytes**********
			//NAME           : StringToBytes
			//PURPOSE        : This function converts string to Bytes.
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006  
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
				 int byteCount= Encoding.Unicode.GetByteCount(str);
					byte[] bytes = new byte[(int)byteCount];
					Encoding.Unicode.GetBytes(str, 0, str.Length, bytes, 0);
					return bytes;
		}
	}
}

using System;
using DanaherTM.ProductEngine;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Scripting;
using DanaherTM.Framework.ExceptionHandling;
using Microsoft.Practices.EnterpriseLibrary.Data;
namespace ExtranetPcat
{
	public class XmlHttpHandler
	{
		DanaherTM.ProductEngine.ProductEngineInstance fooPE;
		XmlDocument xmlDom;
		//private string strBrand = "FNet";
		private string productBrand="";
		private DBUtils.DataStructures.AccessLevel securityCode;
		ArrayList productArray=new ArrayList();
		public XmlHttpHandler()
		{
				productBrand=System.Configuration.ConfigurationSettings.AppSettings["Brand"].ToString();
				//Constructor.
		}
		public void GetProductsforAsset(int assetPID)
			//**********GetProductsforAsset**********
			//NAME           : GetProductsforAsset
			//PURPOSE        : a common function to retrive the products
			//PARAMETERS     : AssetPID as long
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
					XmlElement xmlProductelement;
					XmlElement xmlProductIdelement;
					XmlElement xmlProductNameelement;
					XmlElement xmlAssetProducts;
					xmlAssetProducts=xmlDom.CreateElement("AssetProducts");
					xmlDom.DocumentElement.FirstChild.AppendChild(xmlAssetProducts);
					
					try
					{                
									Product objAsset = new Product(assetPID);
									foreach(Product objProduct in objAsset.ParentProducts)
									{
													xmlProductelement               = xmlDom.CreateElement("Product");
													xmlProductIdelement             = xmlDom.CreateElement("ProductId");
													xmlProductIdelement.InnerText   = Convert.ToString(objProduct.ID);
													xmlProductNameelement           = xmlDom.CreateElement("ProductName");
													xmlProductNameelement.InnerText = objProduct.Name;
													xmlProductelement.AppendChild(xmlProductIdelement);
													xmlProductelement.AppendChild(xmlProductNameelement);
													xmlAssetProducts.AppendChild(xmlProductelement);
													xmlProductelement               = null;
													xmlProductIdelement				= null;
													xmlProductNameelement			= null;
									}
					}
					catch (ProductEngineException ex)
					{
											if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
											{   
												//throw;
											}
					}
					catch(Exception ex)
					{
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
										{   
											//throw;
										}
					}
					finally
					{
										xmlProductNameelement=null;
										xmlProductelement=null;
										xmlProductIdelement=null;
										xmlAssetProducts=null;
					}			
		}
		
		public string CreateModifyAsset(bool isClone, string assetPID ,string assetTitle,
		string assetDescription,string assetFileName, string assetFileSize, DateTime assetBeginDate, 
		string[] assetRelatedProducts, string assetLanguage,string assetMode,string assetSubType,string assetOracleNumber,
		string assetAccess,string assetIndustry,string localeIncludeExclude,bool statusAsset,string oldLanguage,
		string oldAssetOracleNumber,long calendarId)
			//**********CreateModifyAsset**********
			//NAME           : CreateModifyAsset
			//PURPOSE        : Function to add asset records to product catalog.
			//PARAMETERS     : AssetPID as long
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//			      Parag.D		 26-04-2006  Added exception handling at some places.
			//***********************************************************************
		{
					IDbTransaction objIDbT = null;
					Database objData = DBUtils.GetDB();
					IDbConnection objIDbConn = objData.GetConnection();
					objIDbConn.Open();
				 //objIDbT = objIDbConn.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted);
					try
					{
						string generatedAssetPID;
						ArrayList localArray;
						Category assetCat;
						//Modified as now only 2 characters are getting passed.
						localArray=GetLocales(assetLanguage);
						string[] Countries;
						string[] accessGroups;
						string[] industryGroups;
						Product objAsset=null;
						char[] Splitter={','};

						accessGroups=assetAccess.Split(Splitter);
						industryGroups=assetIndustry.Split(Splitter);
						Countries=localeIncludeExclude.Split(Splitter);
						
						if (localeIncludeExclude.Trim().StartsWith("1")==true)
						{
									localArray.Clear();
									foreach(string Country in Countries)
									{
												Locales LanguageLocales=new  Locales();
												foreach(Locale LangLocale in LanguageLocales)
												{
															if (LangLocale.LocaleValue.Trim().ToLower().EndsWith(Country.Trim().ToLower()))
																	if (Country!="1" && Country.Trim() !="")
																			localArray.Add(LangLocale.LocaleValue);
												}
									}
						}

						if (!isClone && assetMode=="A")
						{
							//Save Main Product
							objAsset = new Product();
						}
						else
						{
							try
							{
												if (assetPID.Trim()=="")
												{
																if (isClone)
																{
																	xmlDom = new XmlDocument();
																	xmlDom.LoadXml("<ProductId>" + "Unable to find parent Asset with PID=" + assetPID + "</ProductId>");
																	goto cont;
																}
																else
																{
																	objAsset = new Product();
																}
												}
												else
												{
															if (assetPID.Trim()=="-1" || assetPID.Trim()=="0")
															{
																			if (isClone)
																			{
																							xmlDom = new XmlDocument();
																							xmlDom.LoadXml("<ProductId>" + "Unable to find parent Asset with PID=" + assetPID + "</ProductId>");
																							goto cont;
																			}
																			else
																			{
																							objAsset = new Product();
																			}
															}
															else
															{
																			objAsset = new Product(Convert.ToInt32(assetPID));
															}
												}

												if (objAsset==null)
												{
													xmlDom = new XmlDocument();
													xmlDom.LoadXml("<ProductId>" + "Unable to find Asset with PID=" + assetPID + "</ProductId>");
													goto cont;
												}
							}
							catch (ProductEngineException ex)
							{
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
										{   
											//throw;
										}
							}
							catch(Exception ex)
							{
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
										{   
											//throw;
										}
										xmlDom = new XmlDocument();
										xmlDom.LoadXml("<ProductId>" + "Unable to Create-Update a record" + "</ProductId>");
										goto cont;
							}
						}
      //Code which checks if record for this oraclepartnumber exists.
						try
						{
											Product AssetProduct= new Product(assetOracleNumber);
											if (AssetProduct != null)
											{
															if (assetPID.Trim().ToString() !="" && assetPID.Trim().ToString() != "0")
															{
																			if (AssetProduct.ID.ToString()	!= assetPID)
																			{
																							xmlDom = new XmlDocument();
																							xmlDom.LoadXml("<ProductId>" + assetOracleNumber + " item Number already exists.Please modify item number " +
																								" and save the record again."
																								+ "</ProductId>");
																							goto cont;
																			}
															}
															else
															{
																			xmlDom = new XmlDocument();
																			xmlDom.LoadXml("<ProductId>" + assetOracleNumber + " item number already exists.Please modify item number " +
																				" and save the record again."
																				+ "</ProductId>");
																			goto cont;	
															}
												}
						}
						catch(Exception FoundDuplicateOraPartnumber)
						{
											//If record is not found for this oraclepartnumber exception is raised.
											//Instead it should return a null object.
											//This error has already been logged in mantis.
											//There is no function for checking if Oraclepartnumber exists in Productlocalized table.
											//The other option is to loop through each product and it's localizations and check if 
											//oraclpartnumber exist however this is time consuming and tedious process.
											//As this logic is for checking of oraclepartnumber,this exception block has been kept empty.
											//If there is an exception it implies that this oraclepartnumber does not exist.
											//This is the reason behind keeping this exception block empty.
						}

						//If the record is cloned then don't update the product record.
						if (!isClone)
						{
										objAsset.Name= "AMS ASSETID [" + calendarId + "]";
										//objAsset.Name = assetTitle;
										objAsset.ProductType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Asset;
										switch(Convert.ToInt32(assetSubType))
										{   //Application Notes
											case 1047:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.AppNote;
												break;
												//Brochures
											case 1074:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheet;
												break;
												//"Case Studies"
											case 1068:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Case_Study;
												break;
												//Catalogs
											case 1069:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Catalog;
												break;
												//Corporate
											case 1078:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Corporate;
												break;
												//Data sheets
											case 1066:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheet;
												break;
												//"Extended Specifications"
											case 1048:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Extended_Specificication;
												break;
												//Flyers
											case 1072:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Flyer;
												break;
												//Images
											case 1079:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Image;
												break;
												//Manuals
											case 1071:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Manual;
												break;
												//Miscellaneous
											case 1077:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Miscellaneous;
												break;
												//PowerPoint Presentations
											case 1067:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Powerpoint_Presentation;
												break;
												//Product Software
											case 1052:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Software;
												break;
												//Virtual Demos
											case 1053:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.VirtualDemo;
												break;
												//White Papers
											case 1070:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.WhitePaper;
												break;
												//x Special Reports - Non Product
											case 1073:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XSpecialReports_NonProduct;
												break;
												//x White papers - Non Product
											case 1075:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XWhitePapers_NonProduct;
												break;
												//x Application Notes -Non Product
											case 1076:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XAppNotes_NonProduct;
												break;
												//Letters
											case 1080:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Letter;
												break;
												//Webcasts
											case 1051:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Webcast;
												break;
												//User's Guide
											case 1054:
												objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Users_Card;
												break;
							}
							string accessSecurity;
						 //Assign security codes.
							foreach(string accesscode in accessGroups )
							{
										accessSecurity=Convert.ToString(accesscode.Trim());
										if (accessSecurity=="nfre")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.Free);
										}
										if (accessSecurity=="nlite")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.Light);
										}
										if (accessSecurity=="nfull")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.Full);
										}
										if (accessSecurity=="nosv")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.OSV);
										}
										if (accessSecurity=="nisv")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.ISV);
										}
										if (accessSecurity=="nhnt")
										{
											securityCode=(securityCode |DBUtils.DataStructures.AccessLevel.HTN);
										}
										if (accessSecurity=="ndna")
										{
											securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.DNA);
										}
										if (accessSecurity=="npna") 	
										{
										securityCode=(securityCode | DBUtils.DataStructures.AccessLevel.PNA);
										}
							}

							objAsset.AccessLevel =securityCode;
							objAsset.DefaultBrand = new Brand(productBrand);
							//This category is not used anywhere.To be removed .
							//If not passed, gives error while saving the record.
							objAsset.DefaultCategory=new Category("DCCA");
							////
							objAsset.Status=DanaherTM.ProductEngine.Product.eStatus.Active;
							//objAsset.Save(objIDbT);
						}

						if (objAsset != null)
						{
							objAsset.Save(objIDbT);
						}

						//Delete the existing relationships.
						if (assetMode!="A" && assetPID.Trim()!="")
						{
								DeleteAsset(Convert.ToString(objAsset.ID),oldLanguage,assetMode,isClone,false,false,oldAssetOracleNumber,objIDbT,objAsset);
						}

						foreach(string industryCode in industryGroups)
						{
									if (industryCode.Trim()!="")
									{
										 assetCat=new Category(industryCode.Trim());
										 assetCat.Products.Add(objIDbT,objAsset,
											DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
											DateTime.Now,DBUtils.DataStructures.EndDate_Never);
											assetCat.Dispose();
									}
						}
						
						//Save localized product
						int lngCnt;
						string assetCatalog;
						for (lngCnt=0;lngCnt<localArray.Count;lngCnt++)
						{   
							if (localArray[lngCnt].ToString().Trim()!="")
							{
								if (ValidateLocale(localeIncludeExclude,Convert.ToString(localArray[lngCnt]).Trim())==true)
								{
												ProductLocalized pLocal;
												pLocal = new ProductLocalized();
												if (assetDescription.Trim()=="")
													pLocal.ShortDescription = assetTitle;
												else
													pLocal.ShortDescription = assetDescription;
												pLocal.Name = assetTitle;                
												pLocal.Locale = Convert.ToString(localArray[lngCnt].ToString().Trim());
												pLocal.Parent = objAsset;
												pLocal.FileAssett = assetFileName;
												pLocal.FileSize   = assetFileSize;
												
												if (statusAsset==true)
												{
													pLocal.StartDate  = assetBeginDate;
												}
												else
												{
													pLocal.StartDate  = DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never;
												}
												pLocal.EndDate = DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never;
												pLocal.OraclePartNum = assetOracleNumber;
												pLocal.LongDescription = assetDescription;
												pLocal.CMSGUID = "";
												pLocal.ListPrice=0;
												try
												{
															pLocal.Save(objIDbT);
												}
												catch (ProductEngineException ex)
												{
													if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
													{   
														throw;
														
													}
												}
												catch(Exception ex)
												{
															if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
															{   
																throw;
																
															}
												}
												pLocal.Dispose();
												assetCatalog=Convert.ToString(localArray[lngCnt]);
												assetCatalog=assetCatalog.Trim().Substring(assetCatalog.Trim().Length-2,2);
									   //Add catalog relationships.
												try
												{
																Catalog ProductCatalog = new Catalog(productBrand + "-" +  assetCatalog.ToUpper());	
																ProductCatalog.Assets.Add(objIDbT, objAsset,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,
																1,DateTime.Now,DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never);
															 ProductCatalog.Dispose();
												}
												catch (ProductEngineException ex)
												{
																if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
																{   
																			throw;
 															}
												}
												catch(Exception ex)
												{
																if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
																{   
																			throw;
																}
												}
										}
								}
						}

						//Relate the asset to products
						foreach(string pid in assetRelatedProducts)
						{
							if (pid.Trim()!="")
							{
											try
											{
															Product objMasterProduct = new Product(Convert.ToInt32(pid));
															objMasterProduct.Assets.Add(objIDbT, objAsset,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
															DateTime.Now,DBUtils.DataStructures.EndDate_Never);
															objMasterProduct.Dispose();
														}
														catch (ProductEngineException ex)
														{
																if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
																{   
																	throw;
																}
														}
														catch(Exception ex)
														{
																	if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
																	{   
																		throw;
																	}
														}
										}
						}
					 //Commit the data.
						if (objIDbT!=null)
						{
								objIDbT.Commit();
								objIDbT.Dispose();
								objIDbConn.Close();
								objIDbConn.Dispose();
								objData = null;
						}

						xmlDom =new XmlDocument();
						generatedAssetPID=Convert.ToString(objAsset.ID);
						xmlDom.LoadXml("<ProductId>" + generatedAssetPID + "</ProductId>");
						generatedAssetPID=xmlDom.InnerXml;
						xmlDom=null;
						return(generatedAssetPID);
					cont:
						{
								if (objIDbT !=null)
								{
									objIDbT.Rollback();
									objIDbT.Dispose();
									objIDbConn.Close();
									objIDbConn.Dispose();
									objData = null;
								}
								generatedAssetPID=xmlDom.InnerXml;
								xmlDom=null;
								return(generatedAssetPID);
						}
					}
					catch (ProductEngineException ex)
					{
									if (objIDbT != null)
									{
												objIDbT.Rollback();
												objIDbT.Dispose();
												objIDbConn.Close();
												objIDbConn.Dispose();
												objData = null;
									}

									if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
									{   
										//throw;
									}
									
									string returnData;
									xmlDom = new XmlDocument();
									xmlDom.LoadXml("<ProductId>" + ex.Message + "</ProductId>");
									returnData = xmlDom.InnerXml;
									xmlDom = null;
									return(returnData);
					}
					catch(Exception ex)
						{
									if (objIDbT!= null)
									{
												objIDbT.Rollback();
												objIDbT.Dispose();
												objIDbConn.Close();
												objIDbConn.Dispose();
												objData = null;
									}

									if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
									{   
										//throw;
									}
									string returnData;
									xmlDom = new XmlDocument();
									xmlDom.LoadXml("<ProductId>" + ex.Message + "</ProductId>");
									returnData = xmlDom.InnerXml;
									xmlDom = null;
									return(returnData);
						}
		}
		
		public string EnemerateProducts(int assetPID)
			//**********EnemerateProducts**********
			//NAME           : EnemerateProducts
			//PURPOSE        : Function to get asset records from product catalog.
			//PARAMETERS     : AssetPID as long
			//RETURN VALUE   : string(products in xml format)
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
			string productList;
			XmlElement xmlProductelement;
			XmlElement xmlProductIdelement;
			XmlElement xmlProductNameelement;
			XmlElement xmlAllProducts;
			
			xmlDom =new XmlDocument();
			xmlDom.LoadXml("<Info><Products></Products><Categories></Categories></Info>");
			try
			{
								xmlAllProducts=xmlDom.CreateElement("AllProducts");
								xmlDom.DocumentElement.FirstChild.AppendChild(xmlAllProducts);
								
								Products FnetProducts;
								ProductEngineInstance objPE = new ProductEngineInstance();
								FnetProducts=objPE.GetProducts(DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Mainframe);
								
								foreach(Product ProductInCategory in FnetProducts)
								{   
													xmlProductelement=xmlDom.CreateElement("Product");
													xmlProductIdelement=xmlDom.CreateElement("ProductId");
													xmlProductIdelement.InnerText=Convert.ToString(ProductInCategory.ID);
													xmlProductNameelement=xmlDom.CreateElement("ProductName");
													xmlProductNameelement.InnerText=ProductInCategory.Name;
													xmlProductelement.AppendChild(xmlProductIdelement);
													xmlProductelement.AppendChild(xmlProductNameelement);
													xmlAllProducts.AppendChild(xmlProductelement);
													xmlProductelement = null;
													xmlProductIdelement = null;
													xmlProductNameelement = null;
													productArray.Add(ProductInCategory.ID);
									}
								xmlAllProducts=null;

								if (assetPID!=0 && assetPID!=-1)
								{
											GetProductsforAsset(assetPID);
								}
								GetTopLevelCategories();
								if (assetPID!=0 && assetPID!=-1)
								{
											GetCategoriesforAsset(assetPID);
								}	
								productList=xmlDom.InnerXml;

								xmlDom=null;
								return productList;
			}
			catch (ProductEngineException ex)
			{
								if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
								{   
									//throw;
								}
								string ErrorDescription;
								xmlDom = new XmlDocument();
								xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
								ErrorDescription=xmlDom.InnerXml;
								xmlDom=null;
								return(ErrorDescription);
			}
			catch(Exception ex)
			{
								string ErrorDescription;
								if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
								{   
									//throw;
								}
								xmlDom = new XmlDocument();
								xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
								ErrorDescription=xmlDom.InnerXml;
								xmlDom=null;
								return(ErrorDescription);
			}
		}
		private ArrayList GetLocales(string strLanguage)
			//**********GetLocales**************************************************
			//NAME           : GetLocales
			//PURPOSE        : Function to get Locales from language.
			//PARAMETERS     : Langauge as string
			//RETURN VALUE   : arraylist(containing localevalue)
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
						Locales LanguageLocales=new  Locales(strLanguage);
						ArrayList LocaleArray=new ArrayList();
						foreach(Locale LangLocale in LanguageLocales)
						{
							LocaleArray.Add(LangLocale.LocaleValue);
						}
						return LocaleArray;
		}
		public bool DeleteAsset(string AssetPID,string assetLanguage,string Mode,bool isClone,bool deleteAll,
		bool setRelationship,string ItemNumber, IDbTransaction Trans,Product Asset)
				//**********DeleteAsset**************************************************
				//NAME           : DeleteAsset
				//PURPOSE        : Function to delete assets from product catalog.
				//PARAMETERS     : AssetPID,Language,Mode(Add,Update,Delete),isClone,DeleteAll,setRelationship,ItemNumber,Trans
				//RETURN VALUE   : boolean
				//USAGE		     : 
				//CREATED ON	 : 11-02-2006 
				//CHANGE HISTORY :Auth        	 Date	   	 Description
				//***********************************************************************
				//SP1**
				//Add one more parameter which differentiates the normal delete from do not set pcat rel delete.
				//Last parameter is used for transaction R&D.Will get removed once transaction issue is solved.
		{
			try
			{
				Product objAsset;
				string AssetCatalog;
				if (Asset == null)
				{
							objAsset = new Product(Convert.ToInt32(AssetPID));  
				}
				else
				{
						objAsset = Asset;
				}

				ArrayList LocalArray;
				LocalArray=GetLocales(assetLanguage);
				
				try
				{ 
					if ((Mode=="U") || (deleteAll==true))
					{
						foreach(Product Assetparent in  objAsset.ParentProducts)
						{
							Assetparent.Assets.Remove(Trans,objAsset);
						}
					}
				}
				catch (ProductEngineException ex)
				{
					if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
					{   
						throw;
					}
				}
				catch(Exception ex)
				{
					if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
					{   
						throw;
					}
				}
				//Declare a product object 
				//get the localization object
				//delete the localizations
				//at the same time retrive the catalog and remove the relationship with catalogs for the product
				int iLocalizedCount;

				if(setRelationship==true)
				{
							for(iLocalizedCount=objAsset.Localizations.Count-1;iLocalizedCount >=0;iLocalizedCount--)
							{
										AssetCatalog=Convert.ToString(objAsset.Localizations[iLocalizedCount].Locale);
										AssetCatalog=AssetCatalog.Trim().Substring(AssetCatalog.Trim().Length-2,2);
										Catalog ProductCatalog = new Catalog(productBrand + "-" + AssetCatalog.ToUpper());
										//ProductCatalog.Assets.Remove(objAsset);
										ProductCatalog.Assets.Remove(Trans,objAsset);
										ProductLocalized ProdLocal=new ProductLocalized(objAsset.Localizations[iLocalizedCount].ID);
										//Commented/ Added By Nikita Jain on 20th feb 2007
										//To resolve issues/ optimize code due to new version of ProductEngine dll
										//Old code
										///ProdLocal.Delete();
										//End old code
										ProdLocal.Delete(Trans);
										//End on 20th Feb 2007
							}
				}
				else
				{
					for(iLocalizedCount=objAsset.Localizations.Count-1;iLocalizedCount >=0;iLocalizedCount--)
					{
						try
						{
							ProductLocalized ProdLocal=new ProductLocalized(objAsset.Localizations[iLocalizedCount].ID);
							if (ProdLocal.OraclePartNum.Trim().ToString()== ItemNumber.Trim().ToString())
							{
								
								AssetCatalog=Convert.ToString(objAsset.Localizations[iLocalizedCount].Locale);
								AssetCatalog=AssetCatalog.Trim().Substring(AssetCatalog.Trim().Length-2,2);
								Catalog ProductCatalog = new Catalog(productBrand + "-" +  AssetCatalog.ToUpper());
								ProductCatalog.Assets.Remove(Trans,objAsset);
								ProdLocal.Delete(Trans);
							}
						}
						catch (ProductEngineException ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
							{   
								throw;
							}
						}
						catch(Exception ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
								throw;
							}
						}
					}
					}
				try
				{ 
					if ((Mode=="U") || (deleteAll==true))
					{
						ArrayList AssetCategories = new ArrayList(objAsset.Categories.Count);
						foreach  (Category Cat in objAsset.Categories)
						{
							AssetCategories.Add(Cat.ID);
						}
						for(int CatCount=0;CatCount<AssetCategories.Count;CatCount++)
						{
							if(AssetCategories[CatCount].ToString()!="")
							{
										Category Cat = new Category(Convert.ToInt32(AssetCategories[CatCount]));
										Cat.Products.Remove(Trans,objAsset);
							}
						}
					}
				}
				catch (ProductEngineException ex)
				{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
							{   
								throw;
							}
				}
				catch(Exception ex)
				{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
										throw;
							}
				}
				//SP1**
				//Will remain same.
				if (((Mode!="U") && (Mode !="A")) || deleteAll==true)
				{
					if(isClone==false)
					{
						objAsset.Delete(Trans);
					}
				}
				return true;
			}
			catch (ProductEngineException ex)
			{
				if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
				{   
					throw;
				}
				return false;
			}
			catch(Exception ex)
			{
				//Code for error handling
				if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
				{   
					throw;
				}
				return false;
			}
		}
		
		public void GetTopLevelCategories()
			//**********GetTopLevelCategories**************************************************
			//NAME           : GetTopLevelCategories
			//PURPOSE        : Function to get the top level categories for an asset
			//PARAMETERS     : 
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
			XmlElement XmlCatelement;
			XmlElement XmlCatIdelement;
			XmlElement XmlCatNameelement;
			XmlElement XmlCat;

			fooPE = new ProductEngineInstance();
			XmlCat=xmlDom.CreateElement("AllCategories");
			xmlDom.DocumentElement.LastChild.AppendChild(XmlCat);

			foreach(Category oCatFirst in fooPE.GetCategories(DBUtils.DataStructures.PostingTypes.Industry)) // .Categories)
			{   
							XmlCatelement=xmlDom.CreateElement("Category");
							XmlCatIdelement=xmlDom.CreateElement("CategoryId");
							XmlCatIdelement.InnerText=Convert.ToString(oCatFirst.Code);
							XmlCatNameelement=xmlDom.CreateElement("CategoryName");
							XmlCatNameelement.InnerText=oCatFirst.Name;
							XmlCatelement.AppendChild(XmlCatIdelement);
							XmlCatelement.AppendChild(XmlCatNameelement);
							XmlCat.AppendChild(XmlCatelement);
							XmlCatelement=null;
							XmlCatIdelement=null;
							XmlCatNameelement=null;
			}
		}
		public void GetCategoriesforAsset(long assetPID)
			//**********GetCategoriesforAsset**************************************************
			//NAME           : GetCategoriesforAsset
			//PURPOSE        : Function to get the saved categories for an asset
			//PARAMETERS     : s
			//RETURN VALUE   : void
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//*********************************************************************************
		{
			Product objAsset;
			XmlElement xmlCatelement;
			XmlElement xmlCatIdelement;
			XmlElement xmlCatNameelement;
			XmlElement xmlAssetCategories;

			xmlAssetCategories=xmlDom.CreateElement("AssetCategories");
			xmlDom.DocumentElement.LastChild.AppendChild(xmlAssetCategories);
			objAsset = new Product(Convert.ToInt32(assetPID));
			try
			{
				foreach (Category objCategory  in objAsset.Categories)
				{
					xmlCatelement=xmlDom.CreateElement("Category");
					xmlCatIdelement=xmlDom.CreateElement("CategoryId");
					xmlCatIdelement.InnerText=Convert.ToString(objCategory.Code);
					xmlCatNameelement=xmlDom.CreateElement("ProductName");
					xmlCatNameelement.InnerText=objCategory.Name;
					xmlCatelement.AppendChild(xmlCatIdelement);
					xmlCatelement.AppendChild(xmlCatNameelement);
					xmlAssetCategories.AppendChild(xmlCatelement);
					xmlCatelement     = null;
					xmlCatIdelement   = null;
					xmlCatNameelement = null;
				}
			}
			catch (ProductEngineException ex)
			{
				if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
				{   
					//throw;
				}
			}
			catch(Exception ex)
			{
				if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
				{   
					//throw;
				}			
			}
			xmlCatNameelement=null;
			xmlCatelement=null;
			xmlCatIdelement=null;
			xmlAssetCategories=null;
			objAsset.Dispose();
		}
		
		private bool ValidateLocale(string includeExclude,string Locale) 
			//**********ValidateLocale**************************************************
			//NAME           : ValidateLocale
			//PURPOSE        : Function that decides whether to add a record for the current locale
			// in productlocalized table or not.
			//PARAMETERS     : IncludeExclude,Locale
			//RETURN VALUE   : boolean
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{  
			string[] Countries;
			char[] Splitter={','};
			Countries=includeExclude.Split(Splitter);
			bool found=false;
			bool countryFound=false;
			if ((includeExclude.Trim()=="none") || (includeExclude.Trim().StartsWith("1")==true)){return true;}
			foreach(string Country in Countries)
			{
				if (Country!="")
				{
					if (Locale.ToUpper().EndsWith(Country.ToUpper())==true)
					{
						countryFound=true;
						if (includeExclude.Trim().StartsWith("0")==true)
						{
							found = false;
						}
					}
				}
			}
			if (countryFound==false)
			{
				if (includeExclude.Trim().StartsWith("0")==true){found=true;}
			}
			return found;
		}
		
		public string ValidateCatalogsLocales(string localeIncludeExclude,string assetLanguage,string assetSubType
		,string assetPID,string oraclePartNumber)
			//**********ValidateCatalogsLocales**********
			//NAME           : ValidateCatalogsLocales
			//PURPOSE        : Function for checking if the locales and catalogs are present for the selected language.
			//PARAMETERS     : string localeIncludeExclude,string assetLanguage
			//RETURN VALUE   : string
			//USAGE		     : 
			//CREATED ON	 : 26-04-2006 
			//CHANGE HISTORY :Auth        	 Date	   	 Description
			//***********************************************************************
		{
			ArrayList localArray;
			int lngCnt;
			bool Localefound=true;
			string assetCatalog;
			string localeResult;
			string[] Countries;
			char[] Splitter={','};
			DBUtils.DataStructures.ProductSubTypes prodSubtype = DBUtils.DataStructures.ProductSubTypes.None ;
			int iLocalizedCount;
			switch(Convert.ToInt32(assetSubType))
			{   //Application Notes
				case 1047:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.AppNote;
					break;
					//Brochures
				case 1074:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheet;
					break;
					//"Case Studies"
				case 1068:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Case_Study;
					break;
					//Catalogs
				case 1069:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Catalog;
					break;
					//Corporate
				case 1078:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Corporate;
					break;
					//Data sheets
				case 1066:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheet;
					break;
				case 1048:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Extended_Specificication;
					break;
					//Flyers
				case 1072:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Flyer;
					break;
					//Images
				case 1079:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Image;
					break;
					//Manuals
				case 1071:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Manual;
					break;
					//Miscellaneous
				case 1077:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Miscellaneous;
					break;
					//PowerPoint Presentations
				case 1067:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Powerpoint_Presentation;
					//End on 26th Feb 2007
					break;
					//Product Software
				case 1052:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Software;
					break;
					//Virtual Demos
				case 1053:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.VirtualDemo;
					//End on 26th Feb 2007
					break;
					//White Papers
				case 1070:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.WhitePaper;
					break;
					//x Special Reports - Non Product
				case 1073:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XSpecialReports_NonProduct;
					break;
					//x White papers - Non Product
				case 1075:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XWhitePapers_NonProduct;
					break;
					//x Application Notes -Non Product
				case 1076:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XAppNotes_NonProduct;
					break;
					//Letters
				case 1080:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Letter;
					//End on 26th Feb 2007
					break;
					//Webcasts
				case 1051:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Webcast;
					break;
					//User's Guide
				case 1054:
					prodSubtype = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Users_Card;
					break;
			}

			if(prodSubtype==DBUtils.DataStructures.ProductSubTypes.None)
			{
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<Validate>" + "SUBTYPE" + "</Validate>");
				localeResult=xmlDom.InnerXml;
				xmlDom=null;
				return localeResult;	
			}

			ProductLocalized objProdLocal;

			localArray=GetLocales(assetLanguage);
			Countries=localeIncludeExclude.Split(Splitter);
			if (localeIncludeExclude.Trim().StartsWith("1")==true)
			{
				localArray.Clear();
				foreach(string Country in Countries)
				{
					Localefound=false;
					Locales LanguageLocales=new  Locales();

					foreach(Locale LangLocale in LanguageLocales)
					{
						if (LangLocale.LocaleValue.Trim().ToLower().EndsWith(Country.Trim().ToLower()))
							if (Country!="1" && Country.Trim() !="")
							{
								if (assetPID !="")
								{
									Product objAsset=new Product(Convert.ToInt32(assetPID));
									for(iLocalizedCount=objAsset.Localizations.Count-1;iLocalizedCount >=0;iLocalizedCount--)
									{
										try
										{
											objProdLocal = new ProductLocalized(objAsset.Localizations[iLocalizedCount].ID);

											if (oraclePartNumber!=objProdLocal.OraclePartNum)
											{
												if (objAsset.Localizations[iLocalizedCount].Locale==LangLocale.LocaleValue)
												{
													Localefound=false;	
													xmlDom = new XmlDocument();
													objProdLocal = new ProductLocalized(objAsset.Localizations[iLocalizedCount].ID);

													xmlDom.LoadXml("<Validate>" + "LOCALE-Asset has already been added for Country=" + Country + ", Locale=" + LangLocale.LocaleValue  
														+ " having Generic Number as " + objProdLocal.OraclePartNum + ".Please select different country." + "</Validate>");

													localeResult=xmlDom.InnerXml;
													xmlDom=null;
													return localeResult;
												}
											}
										}
										catch (ProductEngineException ex)
										{
													if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
													{   
														//throw;
													}
										}
										catch(Exception ex)
										{
														if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
														{   
															//throw;
														}
										}
									}
								}
								localArray.Add(LangLocale.LocaleValue);
								Localefound=true;
							}
					}
				}
				if (Localefound==false)
				{
					xmlDom = new XmlDocument();
					xmlDom.LoadXml("<Validate>" + "FALSE" + "</Validate>");
					localeResult=xmlDom.InnerXml;
					xmlDom=null;
					return localeResult;
				}
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<Validate>" + "TRUE" + "</Validate>");
				localeResult=xmlDom.InnerXml;
				xmlDom=null;
				return localeResult;
			}
			
			if (localArray.Count > 0)
			{
				for (lngCnt=0;lngCnt<localArray.Count;lngCnt++)
				{   
					if (localArray[lngCnt].ToString().Trim()!="")
					{
						assetCatalog=Convert.ToString(localArray[lngCnt]);
						assetCatalog=assetCatalog.Trim().Substring(assetCatalog.Trim().Length-2,2);
						try
						{
							string strCatalog = productBrand + "-" + assetCatalog;
							fooPE = new ProductEngineInstance();
							foreach(Catalog CatLocale in fooPE.Catalogs)
							{
								if (CatLocale.Name.ToLower() == strCatalog.ToLower())
								{
									if (CatLocale==null)
									{
										xmlDom = new XmlDocument();
										xmlDom.LoadXml("<Validate>" + "FALSE" + "</Validate>");
										localeResult=xmlDom.InnerXml;
										xmlDom=null;
										return localeResult;
									}
									break;
								}
							}
						}
						catch(ProductEngineException ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
							{   
								//throw;
							}
						}
						catch(Exception ex)
						{
							if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
							{   
								//throw;
							}
							xmlDom = new XmlDocument();
							xmlDom.LoadXml("<Validate>" + "FALSE" + "</Validate>");
							localeResult=xmlDom.InnerXml;
							xmlDom=null;
							return localeResult;
						}
					}
				}
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<Validate>" + "TRUE" + "</Validate>");
				localeResult=xmlDom.InnerXml;
				xmlDom=null;
				return localeResult;
			}
			else
			{
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<Validate>" + "FALSE" + "</Validate>");
				localeResult=xmlDom.InnerXml;
				xmlDom=null;
				return localeResult;
			}
		}
		public string GetProducts()
			//**********SetRelationships**************************************************
			//NAME           : GetProducts
			//PURPOSE        : Function to get products which are not linked.
			//PARAMETERS     : 
			//RETURN VALUE   : 
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	      Description
			//***********************************************************************
		{
			string ProductList;
			XmlElement xmlProductelement;
			XmlElement xmlProductIdelement;
			XmlElement xmlProductNameelement;
			XmlElement xmlLinkedProducts;
			XmlElement xmlNotLinkedProducts;
			Products linkedAssets;
			bool notRelated=false;
			
			xmlDom =new XmlDocument();
			xmlDom.LoadXml("<Info><Products></Products></Info>");
			try
			{
				xmlLinkedProducts=xmlDom.CreateElement("LinkedProducts");
				xmlDom.DocumentElement.FirstChild.AppendChild(xmlLinkedProducts);
				
				xmlNotLinkedProducts=xmlDom.CreateElement("NotLinkedProducts");
				xmlDom.DocumentElement.FirstChild.AppendChild(xmlNotLinkedProducts);
				Products FnetProducts;
				ProductEngineInstance objPE = new ProductEngineInstance();
				FnetProducts=objPE.GetProducts(DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Mainframe);
				
				foreach(Product ProductInCategory in FnetProducts)
				{   
					if ( ProductInCategory.ProductType == DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Mainframe)
					{   
						linkedAssets=ProductInCategory.ProductsByRelationship(DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Asset.ToString());

							if (linkedAssets != null)
							{
											if (linkedAssets.Count > 0)
											{
														xmlProductelement=xmlDom.CreateElement("Product");
														xmlProductIdelement=xmlDom.CreateElement("ProductId");
														xmlProductIdelement.InnerText=Convert.ToString(ProductInCategory.ID);
														xmlProductNameelement=xmlDom.CreateElement("ProductName");
														xmlProductNameelement.InnerText=ProductInCategory.Name;
														xmlProductelement.AppendChild(xmlProductIdelement);
														xmlProductelement.AppendChild(xmlProductNameelement);
														xmlLinkedProducts.AppendChild(xmlProductelement);
														xmlProductelement = null;
														xmlProductIdelement = null;
														xmlProductNameelement = null;
														notRelated = false ;
											}
											else
											{
												notRelated=true;
											}
							}
							else
							{
								notRelated=true;
							}
						if (notRelated)
						{
										xmlProductelement=xmlDom.CreateElement("Product");
										xmlProductIdelement=xmlDom.CreateElement("ProductId");
										xmlProductIdelement.InnerText=Convert.ToString(ProductInCategory.ID);
										xmlProductNameelement=xmlDom.CreateElement("ProductName");
										xmlProductNameelement.InnerText=ProductInCategory.Name;
										xmlProductelement.AppendChild(xmlProductIdelement);
										xmlProductelement.AppendChild(xmlProductNameelement);
										xmlNotLinkedProducts.AppendChild(xmlProductelement);
										xmlProductelement = null;
										xmlProductIdelement = null;
										xmlProductNameelement = null;
										notRelated = false ;
						}
					}
				}
				xmlLinkedProducts=null;
				xmlNotLinkedProducts=null;
				ProductList=xmlDom.InnerXml;
				xmlDom=null;
				return ProductList;
			}
			catch (ProductEngineException ex)
			{
						if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
						{   
							//throw;
						}
						string ErrorDescription;
						xmlDom = new XmlDocument();
						xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
						ErrorDescription=xmlDom.InnerXml;
						xmlDom=null;
						return(ErrorDescription);
			}
			catch(Exception ex)
			{
						string ErrorDescription;
						if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
						{   
							//throw;
						}
						xmlDom = new XmlDocument();
						xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
						ErrorDescription=xmlDom.InnerXml;
						xmlDom=null;
						return(ErrorDescription);
			}
		}
		public string GetProductAssets(string productId)
			//**********SetRelationships**************************************************
			//NAME           : GetProductAssets
			//PURPOSE        : Function to get assets for a product.
			//PARAMETERS     : 
			//RETURN VALUE   : productId
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	      Description
			//***********************************************************************
		{
			try
			{
							Product ProductforAssets = new Product(Convert.ToInt32(productId));
							Products ProductAssets;
							string strAssets="";
							ProductAssets = ProductforAssets.ProductsByRelationship(DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Asset.ToString());

							for (int iAssets=0;iAssets<ProductAssets.Count;iAssets++)
							{
								strAssets += ProductAssets[iAssets].ID + ",";
							}
							xmlDom = new XmlDocument();
							if (strAssets!="")
							{
								xmlDom.LoadXml("<Assets>" + strAssets.Substring(0,strAssets.Length-1) + "</Assets>");
							}
							else
							{
								xmlDom.LoadXml("<Assets>" + "NF"  + "</Assets>");
							}
							strAssets=xmlDom.InnerXml;
							xmlDom=null;
							ProductforAssets =null;
							ProductAssets =null;
							return strAssets;	
			}
			catch (ProductEngineException ex)
			{
						if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
						{   
							//throw;
						}
						string ErrorDescription;
						xmlDom = new XmlDocument();
						xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
						ErrorDescription=xmlDom.InnerXml;
						xmlDom=null;
						return(ErrorDescription);
			}
			catch(Exception ex)
			{
						string ErrorDescription;
						if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
						{   
							//throw;
						}
						xmlDom = new XmlDocument();
						xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
						ErrorDescription=xmlDom.InnerXml;
						xmlDom=null;
						return(ErrorDescription);
			}
		}
		public string SetRelationships(string productId,string assetId)
			//**********SetRelationships**************************************************
			//NAME           : SetRelationships
			//PURPOSE        : Function to set product- asset relationships.
			//PARAMETERS     : 
			//RETURN VALUE   : productId,assetId
			//USAGE		     : 
			//CREATED ON	 : 11-02-2006 
			//CHANGE HISTORY :Auth        	 Date	   	      Description
			//                Zensar        20th feb 2007  Added logic for transaction
			//***********************************************************************
				{
								string[] assetIds;
								char[] Splitter={','};
								assetIds=assetId.Split(Splitter);
								string setRel="True";
								string result;
								
								//Commented/ Added By Nikita Jain on 20th feb 2007
			     //Does not matter even if transactions are not used here.
			     //Need to check.
								//To resolve issues/ optimize code due to new version of ProductEngine dll
								IDbTransaction objIDbT = null;
								Database objData = DBUtils.GetDB();
								IDbConnection objIDbConn = objData.GetConnection();
								//								objIDbConn.Open();
								//								objIDbT = objIDbConn.BeginTransaction();
								//End on 20th Feb 2007
								//To test
								//objIDbT = null;

								try
								{
												Product objMasterProduct = new Product(Convert.ToInt32(productId));
												for(int assetCount=0;assetCount<assetIds.Length;assetCount++)
												{
														Product objAsset = new Product(Convert.ToInt32(assetIds[assetCount]));
														objMasterProduct.Assets.Add(objIDbT,objAsset,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
														DateTime.Now,DBUtils.DataStructures.EndDate_Never);
														objAsset =null;
												}
												if (objIDbT != null)
												{			
													objIDbT.Commit();
													objIDbT.Dispose();
													objIDbConn.Close();
													objIDbConn.Dispose();
												}
								}
								catch (ProductEngineException ex)
								{
									if (objIDbT != null)
									{
											objIDbT.Rollback();
											objIDbT.Dispose();
											objIDbConn.Close();
											objIDbConn.Dispose();
									}
									if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine, ex))
											{   
												//throw;
												setRel="false";
											}
								}
								catch(Exception ex)
								{
										if (objIDbT != null)
										{
													objIDbT.Rollback();
											  objIDbT.Dispose();
													objIDbConn.Close();
													objIDbConn.Dispose();
										}
										if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages, ex))
											{   
												//throw;
													setRel="false";
											}
								}
								xmlDom =new XmlDocument();
								xmlDom.LoadXml("<Result>" + setRel + "</Result>");
								result=xmlDom.InnerXml;
								xmlDom=null;
								return(result);
				}
	}
}

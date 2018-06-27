#region CVS Data
/*
	* CVS Data
	* ----------------------------------------------------------------------------
	* $Source: ,v $
	* $Author: pdeshpan $
	* $Revision: 1.2 $
	* $Date: 2006/02/11 20:00:09 $
	* $Log: ,v $
	* ----------------------------------------------------------------------------
*/
#endregion
using System;
using DanaherTM.ProductEngine;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Scripting;
namespace ProductEngine_TEST
{
	public class XmlHttpHandler
	{
		DanaherTM.ProductEngine.ProductEngineInstance fooPE;
		XmlDocument xmlDom;
		private string strLocale = "en-us";
		private string strCatalog = "FNET-US";
		private string strBrand = "FNet";
		private  enum SecurityLevel
		{
			none=0,
			nfre=1,
			nlite=2,
			nfull=4,
			nisv=8,
			nosv=16,
			ndna=32,
			nhnt=64,
			npna=128
		}
		private SecurityLevel securityCode;
		ArrayList productArray=new ArrayList();
		public XmlHttpHandler()
		{
			fooPE = new ProductEngineInstance(strLocale,DateTime.Now);
		}
		
		//**********GetProductsforAsset**********
		//NAME           : GetProductsforAsset
		//PURPOSE        : a common function to retrive the products
		//PARAMETERS     : AssetPID as long
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************

		public void GetProductsforAsset(long AssetPID)
		{
			Product objAsset;
			XmlElement xmlProductelement;
			XmlElement xmlProductIdelement;
			XmlElement xmlProductNameelement;
			XmlElement xmlAssetProducts;
			xmlAssetProducts=xmlDom.CreateElement("AssetProducts");
			xmlDom.DocumentElement.FirstChild.AppendChild(xmlAssetProducts);
			objAsset = new Product(Convert.ToInt32(AssetPID));
			try
			{
				foreach (Product objProduct  in objAsset.ParentProducts)
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
					xmlProductIdelement   = null;
					xmlProductNameelement = null;
				}
			}
			catch(Exception Ex)
			{
				//Do nothing
			}
			xmlProductNameelement=null;
			xmlProductelement=null;
			xmlProductIdelement=null;
			xmlAssetProducts=null;
			objAsset.Dispose();
		}
		//**********CreateModifyAsset**********
		//NAME           : CreateModifyAsset
		//PURPOSE        : Function to add asset records to product catalog.
		//PARAMETERS     : AssetPID as long
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public string CreateModifyAsset(bool isClone, string assetPID ,string assetTitle,
			string assetDescription,string assetFileName, string assetFileSize, DateTime assetBeginDate, 
			string[] assetRelatedProducts, string assetLanguage,string assetMode,string assetSubType,string assetOracleNumber,
			string assetAccess,string assetIndustry,string localeIncludeExclude)
			//,bool InsertYN
		{
			try
			{
				//string exp;
				string generatedAssetPID;
				ArrayList localArray;
				Category assetCat;
				//localArray=GetLocales(assetLanguage.Substring(0,2));
				//Modified as now only 2 characters are getting passed.
				localArray=GetLocales(assetLanguage);
				string[] Countries;
				string[] accessGroups;
				string[] industryGroups;
				Product objAsset;
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
							if (isClone==true)
							{
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
							if (assetPID.Trim()=="-1")
							{
								objAsset = new Product();
							}
							else
							{
								objAsset = new Product(Convert.ToInt32(assetPID));
							}
						}

						if (objAsset==null)
						{
							xmlDom.LoadXml("<ProductId>" + "Unable to find Asset with PID=" + assetPID + "</ProductId>");
							goto cont;
						}
					}
					catch(Exception ex)
					{
						xmlDom.LoadXml("<ProductId>" + "Unable to Create-Update a record" + "</ProductId>");
						goto cont;
					}
				}
				
				if (!isClone)
				{
					objAsset.Name = assetTitle;
					objAsset.ProductType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Asset;
					switch(Convert.ToInt32(assetSubType))
					{   //Application Notes
						case 1047:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.AppNotes;
							break;
							//Brochures
						case 1074:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheets;
							break;
							//"Case Studies"
						case 1068:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Case_Studies;
							break;
							//Catalogs
						case 1069:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Catalogs;
							break;
							//Corporate
						case 1078:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Corporate;
							break;
							//Data sheets
						case 1066:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheets;
							break;
							//"Extended Specifications"
						case 1048:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Extended_Specificications;
							break;
							//Flyers
						case 1072:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Flyers;
							break;
							//Images
						case 1079:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Images;
							break;
							//Manuals
						case 1071:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Manuals;
							break;
							//Miscellaneous
						case 1077:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Miscellaneous;
							break;
							//PowerPoint Presentations
						case 1067:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Powerpoint_Presentations;
							break;
							//Product Software
						case 1052:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Software;
							break;
							//Virtual Demos
						case 1053:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.VirtualDemos;
							break;
							//White Papers
						case 1070:
							objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.WhitePapers;
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
							//                  case 1080:
							//						objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Letters;
							//						break;
							//					//Webcasts
							//					case 1075:
							//						objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XWhitePapers_NonProduct;
							//						break;
							//					//Letters
							//					case 1080:
							//						objAsset.ProductSubType = DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.XAppNotes_NonProduct;
							//						break;
					}
					string accessSecurity;
					securityCode=SecurityLevel.none;
					foreach(string accesscode in accessGroups )
					{
						accessSecurity=Convert.ToString(accesscode.Trim());
						if (accessSecurity=="nfre")
						{
							securityCode=(securityCode | SecurityLevel.nfre);
						}
						if (accessSecurity=="nlite")
						{
							securityCode=(securityCode | SecurityLevel.nlite);
						}
						if (accessSecurity=="nfull")
						{
							securityCode=(securityCode | SecurityLevel.nfull);
						}
						if (accessSecurity=="nosv")
						{
							securityCode=(securityCode | SecurityLevel.nosv);
						}
						if (accessSecurity=="nisv")
						{
							securityCode=(securityCode | SecurityLevel.nisv);
						}
						if (accessSecurity=="nhnt")
						{
							securityCode=(securityCode | SecurityLevel.nhnt);
						}
						if (accessSecurity=="ndna")
						{
							securityCode=(securityCode | SecurityLevel.ndna);
						}
						if (accessSecurity=="npna") 	
						{
							securityCode=(securityCode | SecurityLevel.npna);
						}
					}
					objAsset.AccessLevel =Convert.ToInt32(securityCode);
					objAsset.DefaultBrand = new Brand("Fnet");
					objAsset.DefaultCategory=new Category("DCCA");
					objAsset.Save();
				}
				if (assetMode!="A" && assetPID.Trim()!="")
				{
					DeleteAsset(Convert.ToString(objAsset.ID),assetLanguage,assetMode,isClone,false);
				}
				

				foreach(string industryCode in industryGroups)
				{
				
					if (industryCode.Trim()!="")
					{
					
						assetCat=new Category(industryCode.Trim());
						assetCat.Products.AddRelationship(objAsset,
							DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
							DateTime.Now,DBUtils.DataStructures.EndDate_Never);
						//fnWriteLog("Industry created" + assetCat.Code + "\n" ,false);
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
							pLocal.ParentProduct = objAsset;
							pLocal.FileAssett = assetFileName;
							pLocal.FileSize   = assetFileSize;
							pLocal.StartDate  = assetBeginDate;
							pLocal.EndDate = DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never;
							pLocal.OraclePartNum = assetOracleNumber;
							pLocal.LongDescription = assetDescription;
							pLocal.CMSGUID = "";
							try
							{
								pLocal.Save();
								//fnWriteLog("Localized created" + pLocal.Locale + "\n",false);
							}
							catch(Exception ex)
							{
								Console.WriteLine("Locale not found" + localArray[lngCnt].ToString().Trim()+ " - " + ex.Message);
								fnWriteLog(objAsset.ID + " - " + assetOracleNumber + " - " + ex.Message + "\n",false);
							}
							pLocal.Dispose();

							assetCatalog=Convert.ToString(localArray[lngCnt]);
							assetCatalog=assetCatalog.Trim().Substring(assetCatalog.Trim().Length-2,2);
							try
							{
								Catalog ProductCatalog = new Catalog("Fnet-" +  assetCatalog.ToUpper());								
								ProductCatalog.Assets.AddRelationship(objAsset,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,
									1,DateTime.Now,DanaherTM.ProductEngine.DBUtils.DataStructures.EndDate_Never);
								//fnWriteLog("catalog relationship created" + ProductCatalog.ID + "\n",false);
							}
							catch(Exception ex)
							{
								Console.WriteLine("Unable to add catalog" + assetCatalog);
								fnWriteLog(objAsset.ID + " - " + assetCatalog + " - " + ex.Message + "\n",false);
								//fnWriteLog("Unable to create catalog relationship " + assetCatalog + "\n",false);
							}
						}
					}
				}

				//Related the asset to products
				foreach(string pid in assetRelatedProducts)
				{
					if (pid.Trim()!="")
					{
						try
						{
							Product objMasterProduct = new Product(Convert.ToInt32(pid));
							objMasterProduct.Assets.AddRelationship(objAsset,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
								DateTime.Now,DBUtils.DataStructures.EndDate_Never);
							//fnWriteLog("created product relationship " + objMasterProduct.ID + "\n",false);
							//objAsset.ParentProducts.AddRelationship(objMasterProduct,
							//DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset,1,
							//DateTime.Now,DBUtils.DataStructures.EndDate_Never);
						}
						catch(Exception ex)
						{
							Console.WriteLine("Unable to add product relationship" + pid);
							fnWriteLog(objAsset.ID + " - " + pid + " - " + ex.Message + "\n",false);
						}
					}
				}

				xmlDom =new XmlDocument();
				generatedAssetPID=Convert.ToString(objAsset.ID);
				xmlDom.LoadXml("<ProductId>" + generatedAssetPID + "</ProductId>");
				generatedAssetPID=xmlDom.InnerXml;
				xmlDom=null;
				return(generatedAssetPID);
			cont:
			{
				generatedAssetPID=xmlDom.InnerXml;
				xmlDom=null;
				return(generatedAssetPID);
			}
			}
			catch(Exception ex)
			{
				string returnData;
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<ProductId>" + ex.Message + "</ProductId>");
				returnData = xmlDom.InnerXml;
				xmlDom = null;
				return(returnData);
			}
		}
		//**********EnemerateProducts**********
		//NAME           : EnemerateProducts
		//PURPOSE        : Function to get asset records from product catalog.
		//PARAMETERS     : AssetPID as long
		//RETURN VALUE   : string(products in xml format)
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public string EnemerateProducts(long AssetPID)
		{
			string ProductList;
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
				
				//Catalog Fnetcatalog=new Catalog(strCatalog);
				Brand FnetBrand=new Brand(strBrand);
				Products FnetProducts;
				//FnetProducts = Fnetcatalog.ProductsByType(DanaherTM.ProductEngine.DBUtils.DataStructures.ProductTypes.Mainframe);
				FnetProducts=FnetBrand.Products;					
				foreach(Product ProductInCategory in FnetProducts)
				{
					//if (ProductInCategory.DefaultBrand.ToString().ToUpper()=="FNET")
					//{
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
					//}
				}
				xmlAllProducts=null;

				if (AssetPID!=0 && AssetPID!=-1)
				{
					GetProductsforAsset(AssetPID);
				}
				GetTopLevelCategories();
				if (AssetPID!=0 && AssetPID!=-1)
				{
					GetCategoriesforAsset(AssetPID);
				}	
				ProductList=xmlDom.InnerXml;

				xmlDom=null;
				return ProductList;
			}
			catch(Exception ex)
			{
				string ErrorDescription;
				xmlDom = new XmlDocument();
				xmlDom.LoadXml("<Info>" + ex.Message + "</Info>");
				ErrorDescription=xmlDom.InnerXml;
				xmlDom=null;
				return(ErrorDescription);
			}
		}
		//**********GetLocales**************************************************
		//NAME           : GetLocales
		//PURPOSE        : Function to get Locales from language.
		//PARAMETERS     : Langauge as string
		//RETURN VALUE   : arraylist(containing localevalue)
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
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
		//**********DeleteAsset**************************************************
		//NAME           : DeleteAsset
		//PURPOSE        : Function to delete assets from product catalog.
		//PARAMETERS     : AssetPID,Language,Mode(Add,Update,Delete),isClone,DeleteAll
		//RETURN VALUE   : boolean
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public bool DeleteAsset(string AssetPID,string Language,string Mode,bool isClone,bool DeleteAll)
		{
			try
			{
				Product objAsset;
				int lngCnt;
				string AssetCatalog;

				objAsset = new Product(Convert.ToInt32(AssetPID));
				ArrayList LocalArray;
				//LocalArray=GetLocales(Language.Substring(0,2));
				LocalArray=GetLocales(Language);

				try
				{ 
					if ((Mode=="U") || (DeleteAll==true))
					{
						foreach(Product Assetparent in  objAsset.ParentProducts)
						{
							Assetparent.Assets.Remove(objAsset);
							//objAsset.ParentProducts.Remove(Assetparent,DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset);
							//objAsset.ParentProducts.Remove(
						}
					}
				}
				catch(Exception Ex)
				{
					//Do nothing
				}

				
				for (lngCnt=0;lngCnt<LocalArray.Count;lngCnt++)
				{  
					try
					{
						ProductLocalized LocalProduct=new ProductLocalized(objAsset,Convert.ToString(LocalArray[lngCnt].ToString().Trim()));
						LocalProduct.Delete();
					}
					catch(Exception ex)
					{
						//Do Nothing
					}
				}

				for (lngCnt=0;lngCnt<LocalArray.Count;lngCnt++)
				{  
					try
					{
						AssetCatalog=Convert.ToString(LocalArray[lngCnt]);
						AssetCatalog=AssetCatalog.Trim().Substring(AssetCatalog.Trim().Length-2,2);
						Catalog ProductCatalog = new Catalog("Fnet-" +  AssetCatalog.ToUpper());
						ProductCatalog.Assets.Remove(objAsset);
					}
					catch(Exception ex)
					{
						//Do Nothing
					}
				}
				
				try
				{ 
					if ((Mode=="U") || (DeleteAll==true))
					{
						//						foreach(Category AssetCat in  objAsset.Categories)
						//						{
						//							//objAsset.Categories.Remove(AssetCat);
						//							Deletecategory(AssetCat.ID,objAsset.ID);
						//						}
						objAsset.Categories.Remove(DanaherTM.ProductEngine.DBUtils.DataStructures.RelationshipTypes.Products.Asset);
					}
				}
				catch(Exception Ex)
				{
					//Do nothing
				}

				if (((Mode!="U") && (Mode !="A")) || DeleteAll==true)
					//if(DeleteAll==true)
				{
					objAsset.Delete();
				}
				return true;
			}
			catch(Exception ex)
			{
				//Code for error handling
				return false;
			}
		}
		//**********GetTopLevelCategories**************************************************
		//NAME           : GetTopLevelCategories
		//PURPOSE        : Function to get the top level categories for an asset
		//PARAMETERS     : 
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public void GetTopLevelCategories()
		{
			XmlElement XmlCatelement;
			XmlElement XmlCatIdelement;
			XmlElement XmlCatNameelement;
			XmlElement XmlCat;
			XmlCat=xmlDom.CreateElement("AllCategories");
			xmlDom.DocumentElement.LastChild.AppendChild(XmlCat);
			foreach(Category oCatFirst in fooPE.Catalogs[strCatalog].Categories)
			{   
				if (oCatFirst.PostingType == DanaherTM.ProductEngine.DBUtils.DataStructures.PostingTypes.Industry && oCatFirst.Localized.Name!="")
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
		}
		//**********GetCategoriesforAsset**************************************************
		//NAME           : GetCategoriesforAsset
		//PURPOSE        : Function to get the saved categories for an asset
		//PARAMETERS     : s
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//*********************************************************************************
		public void GetCategoriesforAsset(long AssetPID)
		{
			Product objAsset;
			XmlElement xmlCatelement;
			XmlElement xmlCatIdelement;
			XmlElement xmlCatNameelement;
			XmlElement xmlAssetCategories;

			xmlAssetCategories=xmlDom.CreateElement("AssetCategories");
			xmlDom.DocumentElement.LastChild.AppendChild(xmlAssetCategories);
			objAsset = new Product(Convert.ToInt32(AssetPID));
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
			catch(Exception Ex)
			{
				Exception ex;
			}
		{//Do nothing
		}
			xmlCatNameelement=null;
			xmlCatelement=null;
			xmlCatIdelement=null;
			xmlAssetCategories=null;
			objAsset.Dispose();
		}
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
		private bool ValidateLocale(string IncludeExclude,string Locale) 
		{  
			string[] Countries;
			char[] Splitter={','};
			Countries=IncludeExclude.Split(Splitter);
			bool found=false;
			bool countryFound=false;
			if ((IncludeExclude.Trim()=="none") || (IncludeExclude.Trim().StartsWith("1")==true)){return true;}
			foreach(string Country in Countries)
			{
				if (Country!="")
				{
					if (Locale.ToUpper().EndsWith(Country.ToUpper())==true)
					{
							countryFound=true;
						if (IncludeExclude.Trim().StartsWith("0")==true)
						{
							found =false;
						}
					}
				}
			}
			if (countryFound==false)
			{
				if (IncludeExclude.Trim().StartsWith("0")==true){found=true;}
			}
			return found;
		}
		//**********UploadExcelFile**************************************************
		//NAME           : UploadExcelFile
		//PURPOSE        : Uploading the data from calendar table and excel sheet into product catalog
		//PARAMETERS     : IncludeExclude,Locale
		//RETURN VALUE   : boolean
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public bool UploadExcelFile(string FileName,string Language)
		{
			string sConnectionOleString =  
				"Provider=Microsoft.Jet.OLEDB.4.0;" + 
				"Data Source=" + FileName + ";" + 
				"Extended Properties=Excel 8.0;"; 
			string PID="";
			

			Hashtable PIDtable=new Hashtable();
			
			string sConnectionString =  
			"Initial Catalog=Fluke_Sitewide;Server=evtibg18.tc.fluke.com;UId=marcomweb;pwd=!?wwwProd1";
			
			OleDbDataReader assetOleDataReader=null;
			SqlDataReader assetDataReader;

			string strProducts="";
			string splitCharater;
			splitCharater=",";
			string strMode="A";
			string sql="";
			
			OleDbConnection objOleConn = new OleDbConnection(sConnectionOleString); 
			SqlConnection objConn=new SqlConnection(sConnectionString);
			
			try
			{
				objConn.Open(); 
				objOleConn.Open();
			}
			catch(Exception Ex)
			{
				return false;
			}
			sql="select distinct Title,Category_ID,SubGroups,'' as Industry,calendar.Description,iso2,[File_Name]," +
				" File_Size,BDate,calendar.Code,'' ProductIds,Clone,File_Name_POD,Revision_Code,item_number,calendar.Id" +
				" from calendar,Language where site_id=82 and file_name is not null " +
				" and calendar.language=language.code " +
				" order by calendar.id";

			SqlCommand objCmdSelect =new SqlCommand(sql,objConn);
			
			SqlDataAdapter objAdapter = new  SqlDataAdapter(); 
			objAdapter.SelectCommand = objCmdSelect; 

			try
			{
				assetDataReader = objCmdSelect.ExecuteReader();
			}
			catch(Exception Ex)
			{
				return false;
			}
			
			string[] productArray;
			string assetFilename;
			string assetPODName;
			
			string strcategory="";
			long lngpid=0;
			string clonePID="";

			while(assetDataReader.Read())
			{   
				assetFilename=assetDataReader.GetValue(6).ToString();
				//assetFile=fileSystem.GetFile(assetFilename);
				//assetFilename=assetFilename.Substring(assetFilename.LastIndexOf(@"\")+1);
				assetPODName=assetDataReader.GetValue(12).ToString();
				
				sql="SELECT IndustryID,ProductID FROM [Sheet1$]" +
					" where Revision_Code='" + assetDataReader.GetValue(13).ToString() +
					"' and item_number=" + assetDataReader.GetValue(14).ToString() ;
				
				OleDbCommand objCmdOleSelect =new OleDbCommand(sql, objOleConn);

				OleDbDataAdapter objOleAdapter = new OleDbDataAdapter();
				objOleAdapter.SelectCommand=objCmdOleSelect;
				try
				{
					assetOleDataReader = objCmdOleSelect.ExecuteReader();
				}
				catch(Exception ex)
				{
					Console.Write(ex.Message);
				}
				try
				{
					assetOleDataReader.Read();
					try
					{
						strProducts=assetOleDataReader.GetValue(1).ToString();
						if (strProducts.Length > 5 )
						{
							Console.WriteLine("Stop");
							strProducts = strProducts.Substring(1,0);
						}
					}
					catch(Exception Ex)
					{
						
					}
					productArray=strProducts.Split(splitCharater.ToCharArray()[0]);
					if (assetOleDataReader.GetValue(0).ToString()=="1")
					{strcategory="DCCA";}
					else if (assetOleDataReader.GetValue(0).ToString()=="2")
					{strcategory="INET";}
					else if (assetOleDataReader.GetValue(0).ToString()=="3")
					{strcategory="TELE";}
					clonePID="0";
					strMode="A";
					if(Convert.ToBoolean(assetDataReader.GetValue(11))==true)
					{
						SqlConnection objcloneConn=new SqlConnection(sConnectionString);
						objcloneConn.Open();
						sql="SELECT PID FROM calendar" +
						" where id=" + assetDataReader.GetValue(11).ToString();
						SqlCommand clonecommand=new SqlCommand(sql,objcloneConn);
						SqlDataReader Clonedatareader=null;
						SqlDataAdapter objcloneadapter = new  SqlDataAdapter(); 
						objcloneadapter.SelectCommand = clonecommand; 
						strMode="U";
						try
						{
							Clonedatareader = clonecommand.ExecuteReader();
							Clonedatareader.Read();
							clonePID =Convert.ToString(Clonedatareader.GetValue(0));

						}
						catch(Exception ex)
						{
							clonecommand.Dispose();
							Clonedatareader.Close();
							objcloneConn.Close();
							objcloneConn.Dispose();
						}
						objcloneConn.Close();
						objcloneConn.Dispose();
						Clonedatareader.Close();
					}
					

					PID=CreateModifyAsset(Convert.ToBoolean(assetDataReader.GetValue(11)),clonePID,assetDataReader.GetValue(0).ToString(),
						assetDataReader.GetValue(4).ToString(), assetFilename,assetDataReader.GetValue(7).ToString(),Convert.ToDateTime(assetDataReader.GetValue(8).ToString()),productArray,
						assetDataReader.GetValue(5).ToString(),strMode,assetDataReader.GetValue(1).ToString(),
						assetDataReader.GetValue(14).ToString(),assetDataReader.GetValue(2).ToString(),
						strcategory,"none");
					assetOleDataReader.Close();
					Console.Write (assetDataReader.GetValue(15).ToString() + " - " + PID + "\n");
					//}
					XmlDocument objxml=new XmlDocument();
					XmlNode objcol;
					objxml.LoadXml(PID);
					if (objxml !=null) 
					{
						objcol=objxml.SelectSingleNode("ProductId");
						PID=objcol.InnerText;
					}
					try
					{
						lngpid=Convert.ToInt32(PID);
						PID=Convert.ToString(lngpid);
					}
					catch(Exception ex)
					{
						PID="0";
					}
					
					if (PID!="0")
					{
						PIDtable.Add(assetDataReader.GetValue(15).ToString(),PID);
						SqlConnection objConnUpdate=new SqlConnection(sConnectionString);
						objConnUpdate.Open();
						SqlCommand UpdatePID=new SqlCommand();
						UpdatePID.CommandText = "update calendar set PID=" + PID + 
							" where id=" +  assetDataReader.GetValue(15).ToString() ;
						//+ 
						//"' and item_number='" + assetDataReader.GetValue(14).ToString()+"'";
						UpdatePID.Connection = objConnUpdate;
						UpdatePID.ExecuteNonQuery();
						UpdatePID.Dispose();
						objConnUpdate.Close();
						objConnUpdate.Dispose();
						fnWriteLog(assetDataReader.GetValue(15).ToString()+ "," + PID + "\n",false);
					}
										
				}
				catch(Exception ex)
				{
					Console.WriteLine(ex.Message);
				}
				objCmdOleSelect =null;
				objOleAdapter=null;
				assetOleDataReader.Close();
			}
			
			objConn.Close(); 
			objConn.Dispose();
			assetDataReader = null;
			return true;
		}
		//**********fnWriteLog**************************************************
		//NAME           : fnWriteLog
		//PURPOSE        : for writing the log to the file
		//PARAMETERS     : text to be logged ,create the file for the first time
		//RETURN VALUE   : void
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		private void fnWriteLog(String sLogTxt ,Boolean blnStart)
		{
			
			StreamWriter oSw=null;
			try
			{	
				if (blnStart==true)
				{
					if (System.IO.File.Exists("C:/log.txt"))
					{
						System.IO.File.Delete("C:/log.txt");
						oSw = System.IO.File.CreateText("C:/log.txt");
						oSw.WriteLine(sLogTxt);
					}
				}
				else
				{
					oSw = System.IO.File.AppendText("C:/log.txt");
					oSw.WriteLine(sLogTxt);
				}
				oSw.Close();
			}

			catch(Exception ex)
			{
				Console.WriteLine (ex.Message);
			}
			finally
			{
				
			}
		}
		//**********updateLocalizedtable**************************************************
		//NAME           : updateLocalizedtable
		//PURPOSE        : Updating the date based on status of asset.
		//PARAMETERS     : 
		//RETURN VALUE   : boolean
		//USAGE		     : 
		//CREATED ON	 : 11-02-2006 
		//CHANGE HISTORY :Auth        	 Date	   	 Description
		//***********************************************************************
		public void updateLocalizedtable()
		{	
			string sConnectionPcat =  
			"Initial Catalog=Fluke_Sitewide;Server=evtibg18.tc.fluke.com;UId=marcomweb;pwd=!?wwwProd1";
		
			string sConnectionString =  
			"Initial Catalog=Fluke_Sitewide;Server=evtibg18.tc.fluke.com;UId=marcomweb;pwd=!?wwwProd1";
			
			SqlDataReader assetDataReader=null;

			string strSqlAssets="";
			
			SqlConnection objPcatConn = new SqlConnection(sConnectionPcat); 
			SqlConnection objConn=new SqlConnection(sConnectionString);
			
			try
			{
				objConn.Open(); 
				objPcatConn.Open();
			}
			catch(Exception ex)
			{
				Console.Write("Unable to open a connection - " + ex.Message);
			}
			
			strSqlAssets = "SELECT dbo.Calendar_backup.language,dbo.Calendar_backup.Id,dbo.Calendar_backup.PID,dbo.Calendar_backup.Status AS Asset_Mgr_Status,dbo.Calendar_backup.Bdate " +
			" FROM  dbo.Literature_Items_US INNER JOIN "+
            " dbo.Calendar_backup ON dbo.Literature_Items_US.ITEM = dbo.Calendar_backup.Item_Number AND "  +
            " dbo.Literature_Items_US.REVISION = dbo.Calendar_backup.Revision_Code LEFT OUTER JOIN " +
            " dbo.Lit_Cost_Center ON dbo.Literature_Items_US.COST_CENTER = dbo.Lit_Cost_Center.Cost_Center " +
			" WHERE (dbo.Lit_Cost_Center.Site_ID = 82) AND (dbo.Literature_Items_US.ACTIVE_FLAG = - 1) " +
			" ORDER BY dbo.Literature_Items_US.ITEM ";
			
			SqlCommand objCmdSelect =new SqlCommand(strSqlAssets , objConn); 
			
			SqlDataAdapter objAdapter = new  SqlDataAdapter(); 
			objAdapter.SelectCommand = objCmdSelect; 
			
			try
			{
				assetDataReader = objCmdSelect.ExecuteReader();
			}
			catch(Exception ex)
			{
				Console.Write("Unable to execute the reader");
			}
			
			ArrayList localArray;
			while(assetDataReader.Read())
			{
				localArray=GetLocales(assetDataReader.GetValue(0).ToString().Substring(0,2));
				Product ModifiedProduct=new Product(Convert.ToInt32(assetDataReader.GetValue(2).ToString()));
				foreach(string LangLocale in localArray)
					{   
						try
						{
							ProductLocalized ModifiedLocalProduct=new ProductLocalized(ModifiedProduct,LangLocale);
							if (assetDataReader.GetValue(3).ToString()=="1")
							{
								ModifiedLocalProduct.StartDate=Convert.ToDateTime(assetDataReader.GetValue(4).ToString());
								ModifiedLocalProduct.Save();
							}
						}
						catch(Exception ex)
						{
							//Do nothing.This product may be excluded.
						}
					}
			}
		}
	}
}

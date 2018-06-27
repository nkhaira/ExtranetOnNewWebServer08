using System;
using DanaherTM.ProductEngine;
using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using DataLibrary = Microsoft.Practices.EnterpriseLibrary.Data;
using DanaherTM.Framework.ExceptionHandling;
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;  
using Microsoft.Practices.EnterpriseLibrary.Configuration;
using System.Configuration;


namespace AssetIndexing
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class ListAssets
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			//
			// TODO: Add code to start application here

			//English Assets
           try
           {
                ListAssets_Eng();
                // French Assets
                ListAssets_Fre();
//                //Italian Assets
                ListAssets_Ita();

                ListAssets_Dut();

                ListAssets_Ger();

                ListAssets_Spa();

                ListAssets_Swe();

                ListAssets_Dan();

                ListAssets_Por();

                ListAssets_Nor();

                ListAssets_Jpn();

                ListAssets_Rus();

                ListAssets_Chi();
    

               
            }
            catch(Exception ex)
            {
               if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.WebPages,ex))
               {
                   //throw;
                //   Console.WriteLine("Unable to write files");
               }
                               
            }

            try
            {
                GenerateXmlForAll();
            }
            catch(ProductEngineException ex)
            {
                if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine,ex))
                {
                    //throw;
                    //   Console.WriteLine("Unable to write files");
                }
                               
            }


        
		}

        
       public static void ListAssets_Eng()
        {
           string sAssetServer,sFilePath;
           StreamWriter oListWriter; 
           int iSiteId;

           iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);

           sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
           sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

           DataLibrary.Database oSiteWideDB=null;
           oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
           sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'eng',"+ iSiteId +"");     

           oListWriter=File.CreateText(sFilePath+"AssetList_en.htm");
           oListWriter.WriteLine("<html>");
           oListWriter.WriteLine("<body>");
           while(sqlReader.Read())
           {
               oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
           }

           oListWriter.WriteLine("</body>");
           oListWriter.WriteLine("</html>");
           
           oListWriter.Close(); 
           sqlReader.Close();
            

        }
        

        public static void ListAssets_Fre()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'fre',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_fr.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
        

        }

        public static void ListAssets_Ita()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'ita',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_it.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
        

        }

        public static void ListAssets_Dut()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter; 
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'dut',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_nl.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
        
        }

        public static void ListAssets_Ger()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter; 
            int iSiteId;

            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'ger',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_de.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();

           

        }

        public static void ListAssets_Spa()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'spa',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_es.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();

           

        }

        public static void ListAssets_Swe()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'swe',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_sv.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
         

        }

        public static void ListAssets_Dan()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter; 
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'dan',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_da.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void ListAssets_Por()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'por',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_pt.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void ListAssets_Nor()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'nor',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_no.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void ListAssets_Jpn()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter; 
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'jpn',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_ja.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void ListAssets_Rus()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'rus',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_ru.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void ListAssets_Chi()
        {
            string sAssetServer,sFilePath;
            StreamWriter oListWriter;
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_Asset_Sel 'chi',"+ iSiteId +"");     

            oListWriter=File.CreateText(sFilePath+"AssetList_zh.htm");
            oListWriter.WriteLine("<html>");
            oListWriter.WriteLine("<body>");
            while(sqlReader.Read())
            {
                oListWriter.WriteLine("<a href='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
            }

            oListWriter.WriteLine("</body>");
            oListWriter.WriteLine("</html>");
           
            oListWriter.Close(); 
            sqlReader.Close();
       

        }

        public static void GenerateXmlForAll()
        {
            string sAssetServer,sCategory,sFindItPath,sFilePath,sSecured;
            StreamWriter oListWriter; 
            int iSiteId;

            sAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
            sFindItPath=System.Configuration.ConfigurationSettings.AppSettings["FindItUrl"].Replace("&","&amp;");
            iSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
            sFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            DataLibrary.Database oSiteWideDB=null;
            oSiteWideDB = DataLibrary.DatabaseFactory.CreateDatabase("FlukeSitewide");
        
            SqlDataReader sqlReader;
                      
            
            //sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.StoredProcedure, "FNET_AllAssets_Sel");     
            sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "FNET_AllAssets_Sel "+ iSiteId +""); 


            oListWriter=File.CreateText(sFilePath+"AssetDetails.xml");
          

            oListWriter.WriteLine("<?xml version='1.0' encoding='UTF-8'?>"); 
            oListWriter.WriteLine("<Assets>");
            
            while(sqlReader.Read())
            {
                

                //if(sqlReader["PCAT_ID"].ToString()!="-1")
                if(Convert.ToInt32(sqlReader["PCAT_ID"])>0)
                {
                   Product objProd;
                    try
                    {
                        objProd =  new Product(Convert.ToInt32(sqlReader["PCAT_ID"]));
                    
                        oListWriter.WriteLine(" <Asset");
                        oListWriter.WriteLine(" URI='"+sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>");
                        oListWriter.WriteLine(" 	<Title>");
                        oListWriter.WriteLine(sqlReader["Title"].ToString().Replace("&","&amp;"));
                        oListWriter.WriteLine(" 	</Title>");
                        oListWriter.WriteLine(" 	<CrawlURI>");
                        oListWriter.WriteLine(sAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]);
                        oListWriter.WriteLine(" 	</CrawlURI>");
                        oListWriter.WriteLine(" 	<UserURI>");
                        oListWriter.WriteLine(sFindItPath+sqlReader["Item_Number"]);
                        oListWriter.WriteLine(" 	</UserURI>");
                        if(objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Manual || objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.AppNote || objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.DataSheet || objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.WhitePaper)
                        {
                            sCategory="Technical";
                        }
                        else if(objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.Software)
                        {
                            sCategory="Support";
                        }
                        else if(objProd.ProductSubType==DanaherTM.ProductEngine.DBUtils.DataStructures.ProductSubTypes.VirtualDemo)
                        {
                            sCategory="Products";
                        }
                        else
                        {
                            sCategory="Asset";
                        }
                        oListWriter.WriteLine(" 	<MainCategory>");
                        oListWriter.WriteLine(sCategory);
                        oListWriter.WriteLine(" 	</MainCategory>");

                        if(Convert.ToInt32(sqlReader["SubGroups"].ToString().IndexOf("nfre"))==-1)
                        {
                            sSecured="yes";
                        }
                        else
                        {
                            sSecured="no";
                        }
    //                  
                        oListWriter.WriteLine(" 	<Secured>");
                        oListWriter.WriteLine(sSecured);
                        oListWriter.WriteLine(" 	</Secured>");
                                            
                        oListWriter.WriteLine(" </Asset>");
                    }
                    catch(ProductEngineException ex)
                    {
                        if (ExceptionPolicy.HandleException(DanaherTM.Framework.ExceptionHandling.ExceptionInstance.FlukeNetworks.ProductEngine,ex))
                        {
                            //throw;
                            //   Console.WriteLine("Unable to write files");
                        }
                    }

                   
                }

               
            }

            oListWriter.WriteLine("</Assets>");
                       
            oListWriter.Close(); 
            sqlReader.Close();
       

        }


        
	}
}

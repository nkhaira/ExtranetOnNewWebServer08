using System;


using System.Collections;
using System.Xml;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using DanaherTM.CommonUtilities;
namespace DanaherTM.Ircon.AssetIndexing
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
			
			try
			{
                //English Assets
				ListAssets_Eng();
//				 French Assets
				ListAssets_Fre();
//               Italian Assets
				ListAssets_Dut();
//
				ListAssets_Ger();
//
				ListAssets_Spa();				
//
				ListAssets_Jpn();				
//
				ListAssets_Chi();
               
			}
			catch(Exception ex)
			{
                ExceptionManagement.HandleException(ex, "AssetsIndex");
                
                               
			}
			/* Creates XML file with urls */
			try
			{
				GenerateXmlForAll();
			}
			catch(Exception ex)
			{
                ExceptionManagement.HandleException(ex, "AssetIndex");
                               
			}        
		}
/*********************************************************************************************/
		/// <summary>
		/// Generates XML file listing assets
		/// </summary>
		public static void GenerateXmlForAll()
		{
			string strAssetServer,strCategory,strFindItPath,strFilePath,strSecured;
			StreamWriter oListWriter; 
			int intSiteId;

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			strFindItPath=System.Configuration.ConfigurationSettings.AppSettings["FindItUrl"].Replace("&","&amp;");
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
			
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.StoredProcedure, "FNET_AllAssets_Sel");     
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_AllAssets_Sel "+ intSiteId +""); 
            //Updated --by Ravi
            SqlParameter[] parmArray = new SqlParameter[1];
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);
            
            spSiteId.Value = intSiteId;
            spSiteId.Direction = ParameterDirection.Input;
            
            parmArray[0] = spSiteId;
                        
            sqlReader = (SqlDataReader)oSiteWideDB.ExecuteDataReader("IRCN_AllAssets_Sel",parmArray);   

			oListWriter=File.CreateText(strFilePath+"AssetDetails.xml");
          
 		    oListWriter.WriteLine("<?xml version='1.0' encoding='UTF-8'?>"); 
			oListWriter.WriteLine("<Assets>");
            
			while(sqlReader.Read())
			{
                //if(sqlReader["PCAT_ID"].ToString()!="-1")
                //Updated as Pcat_Id no more used in Ircon -- by Ravi 
				//if(Convert.ToInt32(sqlReader["PCAT_ID"])>0)
				//{
					try
					{
						//objProd =  new Product(Convert.ToInt32(sqlReader["PCAT_ID"]));
                    
						oListWriter.WriteLine(" <Asset");
						oListWriter.WriteLine(" URI='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>");
						oListWriter.WriteLine(" 	<Title>");
						oListWriter.WriteLine(sqlReader["Title"].ToString().Replace("&","&amp;"));
						oListWriter.WriteLine(" 	</Title>");
						oListWriter.WriteLine(" 	<CrawlURI>");
						oListWriter.WriteLine(strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]);
						oListWriter.WriteLine(" 	</CrawlURI>");
						oListWriter.WriteLine(" 	<UserURI>");
						oListWriter.WriteLine(strFindItPath+sqlReader["Item_Number"]);
						oListWriter.WriteLine(" 	</UserURI>");
						oListWriter.WriteLine(" </Asset>");
					}
					catch(Exception ex)
					{
                        Console.WriteLine("Unable to write files" + ex.Message);
                        ExceptionManagement.HandleException(ex, "AssetsIndex");
                              
					}
            //}
			}
			oListWriter.WriteLine("</Assets>");
            oListWriter.Close(); 
			sqlReader.Close();
		}

/*********************************************************************************************/   
	/// <summary>
	/// Lists English language specific assets
	/// </summary>
		public static void ListAssets_Eng()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter; 
			int intSiteId;

			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'eng',"+ intSiteId +"");     
            //Updated --by Ravi
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode",SqlDbType.VarChar,5);   
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);
            
            spLangCode.Value ="eng";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;  
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;  
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel",parmArray); 
    
			oListWriter=File.CreateText(strFilePath+"AssetList_en.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();
            

		}

/*********************************************************************************************/
		/// <summary>
		/// Lists French language specific assets
		/// </summary>
		public static void ListAssets_Fre()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter;
			int intSiteId;
               
			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'fre',"+ intSiteId +"");     
            //Updated --by Ravi
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "fre";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 
			oListWriter=File.CreateText(strFilePath+"AssetList_fr.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();
        

		}
	
/*********************************************************************************************/
		/// <summary>
		/// Lists dutch language specific assets
		/// </summary>
		public static void ListAssets_Dut()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter; 
			int intSiteId;

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
			
			SqlDataReader sqlReader;
            //Updated --by Ravi          
            //sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'dut',"+ intSiteId +"");     
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "dut";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 
			oListWriter=File.CreateText(strFilePath+"AssetList_nl.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();
        
		}

/*********************************************************************************************/
		/// <summary>
		/// Lists German language specific assets
		/// </summary>
		public static void ListAssets_Ger()
		{
            string strAssetServer,strFilePath;
			StreamWriter oListWriter; 
			int intSiteId;

			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
			SqlDataReader sqlReader;
                                  
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'ger',"+ intSiteId +"");     
            //Updated -- by Ravi 
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "ger";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 
			oListWriter=File.CreateText(strFilePath+"AssetList_de.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();

           

		}

/*********************************************************************************************/
		/// <summary>
		/// Lists spanish language specific assets
		/// </summary>
		public static void ListAssets_Spa()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter;
			int intSiteId;

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'spa',"+ intSiteId +"");     
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "spa";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 

			oListWriter=File.CreateText(strFilePath+"AssetList_es.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();

           

		}
	
/*********************************************************************************************/
		/// <summary>
		/// Lists Japanese language specific assets
		/// </summary>
		public static void ListAssets_Jpn()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter; 
			int intSiteId;

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'jpn',"+ intSiteId +"");     
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "jpn";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 

			oListWriter=File.CreateText(strFilePath+"AssetList_ja.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();
       

		}
	
/*********************************************************************************************/
		/// <summary>
		/// Lists Chinese language specific assets
		/// </summary>
		public static void ListAssets_Chi()
		{
			string strAssetServer,strFilePath;
			StreamWriter oListWriter;
			int intSiteId;

			strAssetServer=System.Configuration.ConfigurationSettings.AppSettings["AssetServer"];
			intSiteId=Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings["SiteCodeID"]);
			strFilePath=System.Configuration.ConfigurationSettings.AppSettings["FilePath"];

            Database oSiteWideDB = new Database("FlukeSitewide");
        
			SqlDataReader sqlReader;
                      
            
			//sqlReader =(SqlDataReader)oSiteWideDB.ExecuteReader(System.Data.CommandType.Text, "IRCN_Asset_Sel 'chi',"+ intSiteId +"");     
            SqlParameter[] parmArray = new SqlParameter[2];
            SqlParameter spLangCode = new SqlParameter("@LangCode", SqlDbType.VarChar, 5);
            SqlParameter spSiteId = new SqlParameter("@SiteCode", SqlDbType.Int, 4);

            spLangCode.Value = "chi";
            spSiteId.Value = intSiteId;

            spLangCode.Direction = ParameterDirection.Input;
            spSiteId.Direction = ParameterDirection.Input;

            parmArray[0] = spLangCode;
            parmArray[1] = spSiteId;

            sqlReader = oSiteWideDB.ExecuteDataReader("IRCN_Asset_Sel", parmArray); 

			oListWriter=File.CreateText(strFilePath+"AssetList_zh.htm");
			oListWriter.WriteLine("<html>");
			oListWriter.WriteLine("<body>");
			while(sqlReader.Read())
			{
				oListWriter.WriteLine("<a href='"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"'>"+strAssetServer+"/"+sqlReader["Site_Code"]+"/"+ sqlReader["File_Name"]+"</a><br>");
               
			}

			oListWriter.WriteLine("</body>");
			oListWriter.WriteLine("</html>");
           
			oListWriter.Close(); 
			sqlReader.Close();
       

		}
	
/*********************************************************************************************/
	}//class ListAssets end
}

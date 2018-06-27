using EL = Microsoft.Practices.EnterpriseLibrary;
using DanaherTM.ProductEngine;
using System;
using System.Data.SqlClient;
using ProductEngine_TEST;
namespace ProductEngine_TEST
{
	public class ImportAssets
	{
		[STAThread]
		static void Main(string[] args)
		{ 
			XmlHttpHandler objxml = new XmlHttpHandler();
			string[] langlist= new string[2];
			langlist[0]="and clone=0";
			langlist[1]="and clone!=0";
			objxml.UploadExcelFile(@"C:\productengine_test\AssetData.xls" ,"");
		}
	}
}

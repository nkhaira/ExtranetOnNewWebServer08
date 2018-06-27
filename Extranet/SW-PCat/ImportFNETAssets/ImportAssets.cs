using EL = Microsoft.Practices.EnterpriseLibrary;
using DanaherTM.ProductEngine;
using System;
using System.Data.SqlClient;
namespace ImportFNETAssets
{
	public class ImportAssets
	{
		[STAThread]
		static void Main(string[] args)
		{ 
			XmlHttpHandler objxml = new XmlHttpHandler();
			string[] langlist= new string[2];
			objxml.fnWriteLog("Process started",true);
			objxml.UploadExcelFile();
			objxml.fnWriteLog("Process completed",false);
		}
	}
}

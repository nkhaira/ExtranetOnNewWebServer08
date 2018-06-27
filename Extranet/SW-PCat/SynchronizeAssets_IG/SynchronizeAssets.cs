//using EL = Microsoft.Practices.EnterpriseLibrary;
//using DanaherTM;
//using DanaherTM.ProductEngine;
using System;
using System.Data.SqlClient;
//using Microsoft.VisualBasic;
namespace SynchronizeAssets
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	/// 
	public class SynchronizeAssets
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{ 
			SyncAssets objxml = new SyncAssets();
			objxml.updateLocalizedtable();
		}
	}
#region User defined functions and procedures
#endregion
}

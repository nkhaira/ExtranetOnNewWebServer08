#region Copyright and License Information
/*
-------------------------------------------------------------------------
DNNPortal-Download a Upload / Download Module for DotNetNuke 3.X
DNNPortal-Download is written in C#

DNNPortal-Download is based on the functionality of DNNDownload 
(c) Steve Fabian and Hans-Peter Schelian
which is based on the idea of Steve Fabians GoodDogs Repository
(c) Steve Fabian (http://www.gooddogs.com/dotnetnuke/)

RSS.NET (http://rss-net.sf.net/) 
Copyright © 2002, 2003 George Tsiokos. All Rights Reserved.


DNNPortal-Download Copyright (C) 2005  Hans-Peter Schelian (hp@schelian.de)
http://www.dnnportal.de


This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License as
published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
-------------------------------------------------------------------------
*/
#endregion

#region Namespaces
using System;
using System.Web;
using System.Drawing.Imaging;
using DotNetNuke.Entities.Portals;
using HPS.DNN.Modules.Download.Business;
#endregion

namespace HPS.DNN.Modules.Download
{
	/// <summary>
	/// class MakeThumbnail
	/// </summary>
	public class MakeThumbnail : System.Web.UI.Page
	{

#region Event Handlers
		/// <summary>
		/// Page_Load event
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void Page_Load(object sender, System.EventArgs e)
		{
			string imageType;
			string imageUrl = Request.QueryString["img"];
			int imageHeight = (Convert.ToInt32 (Request.QueryString["h"]));
			int imageWidth = (Convert.ToInt32 (Request.QueryString["w"]));
			imageType = Utils.getImageTypeForSecureFile(imageUrl);
			PortalSettings portalSettings = ((PortalSettings)(HttpContext.Current.Items["PortalSettings"]));
			string pathToImage = portalSettings.HomeDirectory + imageUrl;
			System.Drawing.Image fullSizeImg;
			fullSizeImg = System.Drawing.Image.FromFile(Server.MapPath(pathToImage));
			int fullHeight = fullSizeImg.Height;
			int fullWidth = fullSizeImg.Width;
			if (imageWidth > 0 & imageHeight == 0) 
			{
				imageHeight = ((int)(((imageWidth * fullHeight) / fullWidth)));
			}
			if (imageHeight > 0 & imageWidth == 0) 
			{
				imageWidth = ((int)(((imageHeight * fullWidth) / fullHeight)));
			}
			if (imageHeight == 0 & imageWidth == 0) 
			{
				imageHeight = fullHeight;
				imageWidth = fullWidth;
			}

			System.Drawing.Image thumbNailImage = fullSizeImg.GetThumbnailImage(imageWidth, imageHeight, new System.Drawing.Image.GetThumbnailImageAbort(ThumbnailCallback), IntPtr.Zero);

			if (imageType == ".JPG") 
			{
				Response.ContentType = "image/jpeg";
				thumbNailImage.Save(Response.OutputStream, ImageFormat.Jpeg);
			} 
			else if (imageType == ".GIF") 
			{
				Response.ContentType = "image/gif";
				thumbNailImage.Save(Response.OutputStream, ImageFormat.Gif);
			}
		}

		/// <summary>
		/// Callback method ThumbnailCallback
		/// </summary>
		/// <returns></returns>
		public bool ThumbnailCallback()
		{
			return true;
		}

#endregion

#region Web Form Designer generated code
		/// <summary>
		/// OnInit event
		/// </summary>
		/// <param name="e"></param>
		override protected void OnInit(EventArgs e)
		{
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		///		Required method for Designer support - do not modify
		///		the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.Load += new System.EventHandler(this.Page_Load);
		}
#endregion
		

	}
}

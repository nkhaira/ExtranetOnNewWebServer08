using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Text;

namespace DanaherTM.FlukeNetworks
{
    /// <summary>
    /// HttpHandler for processing asset tool requests
    /// </summary>
    public class AssetAspHttpHandler : DefaultHttpHandler
    {
        public AssetAspHttpHandler()
        {
        }

        #region DefaultHttpHandler overrides
        /// <param name="context">An <see cref="T:System.Web.HttpContext"></see> that provides references to intrinsic server objects used to service HTTP requests.</param>
        /// <param name="callback">The <see cref="T:System.AsyncCallback"></see> to call when the asynchronous method call is complete. If callback is null, the delegate is not called.</param>
        /// <param name="state">Any state data needed to process the request.</param>
        /// <returns>An <see cref="T:System.IAsyncResult"></see> that contains information about the status of the process.</returns>
        public override IAsyncResult BeginProcessRequest(HttpContext context, AsyncCallback callback, object state)
        {
            // perform our actions here. any custom code should get called from this method.
            PreProcessActivities(context);

            // continue with pipeline
            return base.BeginProcessRequest(context, callback, state);
        }

        /// <param name="context">An <see cref="T:System.Web.HttpContext"></see> that provides references to intrinsic server objects used to service HTTP requests.</param>
        public override void ProcessRequest(HttpContext context)
        {
            // perform our actions here. any custom code should get called from this method.
            PreProcessActivities(context);

            // continue with pipeline
            base.ProcessRequest(context);
        }
        #endregion

        /// <summary>
        /// method responsible for all preprocessing activities
        /// </summary>
        /// <param name="context">An <see cref="T:System.Web.HttpContext"></see> that provides references to intrinsic server objects used to service HTTP requests.</param>
        private void PreProcessActivities(HttpContext context)
        {
            Uri targetUrl = context.Request.Url;
            bool checkDomain = false;
            string handlerDomains = ConfigurationManager.AppSettings["HandlerDomain"];
            string[] arrDomains = handlerDomains.Split(';');
            foreach (string domain in arrDomains)
            {
                if (targetUrl.Host == domain)
                {
                    checkDomain = true;
                    break;
                }
            }

            if (checkDomain)
            {
                if (context.Request["document"] != null)
                {
                    string accessIDs = String.Empty;
                    string[] aIDs;
                    string role;
                    bool status = false;
                    string userDetails;
                    string userName;
                    StringBuilder returnUrl = new StringBuilder();
                    System.Web.HttpCookie userCookie = null;
                    MembershipUser currentUser;
                    SqlDataReader drAsset = null;

                    SqlConnection oConn = new SqlConnection(ConfigurationManager.ConnectionStrings["AssetDB"].ConnectionString);
                    oConn.Open();

                    SqlCommand oCmd = new SqlCommand();
                    oCmd.CommandText = "FNET_HTTPHANDLER_GETASSET";
                    oCmd.CommandType = CommandType.StoredProcedure;
                    oCmd.Connection = oConn;

                    SqlParameter docId = new SqlParameter();
                    docId.ParameterName = "@documentID";
                    docId.SqlDbType = SqlDbType.VarChar;
                    docId.Size = 20;
                    docId.Value = context.Request["document"];

                    oCmd.Parameters.Add(docId);

                    drAsset = oCmd.ExecuteReader(CommandBehavior.CloseConnection);

                    if (drAsset.Read())
                    {
                        accessIDs = drAsset["SubGroups"].ToString();
                    }
                    if (accessIDs.Length == 0)
                        return;
                    else
                    {
                        if (accessIDs.ToLower().IndexOf("nfre") > 0)
                            return;
                        else
                        {
                            returnUrl.Append(targetUrl.AbsoluteUri);
                            if (returnUrl.ToString().IndexOf("?") > 0)
                            {
                                if (returnUrl.ToString().ToLower().IndexOf("src") == -1)
                                    returnUrl.Append("&Src=").Append(context.Request["Src"]);
                                if (returnUrl.ToString().ToLower().IndexOf("style") == -1)
                                    returnUrl.Append("&Style=").Append(context.Request["Style"]);
                                if (returnUrl.ToString().ToLower().IndexOf("document") == -1)
                                    returnUrl.Append("&document=").Append(context.Request["document"]);
                            }
                            else
                                returnUrl.Append("?Src=").Append(context.Request["Src"]).Append("&Style=").Append(context.Request["Style"]).Append("&document=").Append(context.Request["document"]);
                            
                            userCookie = context.Request.Cookies["FNetUser"];
                            // check if cookie is set. Cookie is set when user logs in.
                            if (userCookie == null)
                                RedirectToLogin(context, accessIDs, returnUrl);
                            else
                            {
                                System.Text.ASCIIEncoding decode = new System.Text.ASCIIEncoding();
                                userDetails = decode.GetString(Convert.FromBase64String(userCookie.Value));
                                userName = userDetails.Split('|')[0];
                                currentUser = Membership.GetUser(userName);
                                if (currentUser == null)
                                    RedirectToLogin(context, accessIDs, returnUrl);
                                else
                                {
                                    string assetRole = string.Empty;
                                    string[] userRoles = Roles.GetRolesForUser(currentUser.UserName);
                                    aIDs = accessIDs.Split(',');
                                    for (int x = 0; x < aIDs.Length; x++)
                                    {
                                        for (int i = 0; i < userRoles.Length; i++)
                                        {
                                            role = userRoles[i].ToLower();
                                            role = role.Replace("gold_", "");

                                            // make the changes wrt to roles stored in assets DB
                                            if (role == "dista")
                                                role = "dna";
                                            if (role == "hhnt")
                                                role = "hnt";
                                            if (role == "ovina")
                                                role = "pna";

                                            role = "n" + role;

                                            if (role == aIDs[x].Trim().ToLower())
                                            {
                                                status = true;
                                                break;
                                            }
                                        }
                                        if (status == true)
                                            break;
                                    }

                                    if (status == false)
                                    {
                                        if (accessIDs.IndexOf("nfull") > 0)
                                        {
                                            //redirect to Edit Personal Info page
                                            context.Response.Redirect(ConfigurationManager.AppSettings["LoginServer"] + "myAccount/register.htm?returnPage=" + context.Server.UrlEncode(returnUrl.ToString()));
                                        }
                                        else
                                        { 
                                            //redirect to gold info page
                                            context.Response.Redirect(ConfigurationManager.AppSettings["LoginServer"] + "supportAndDownloads/aboutGold/");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This method redirects to Login page depending on the access level of assets
        /// </summary>
        /// <param name="accessIDs">comma seperated access string of asset</param>
        /// <param name="returnUrl">return Url to hit secured assets after successfully login</param>
        private void RedirectToLogin(HttpContext context, string accessIDs, StringBuilder returnUrl)
        {
            if (accessIDs.ToLower().IndexOf("nlite") > 0)
            {
                context.Response.Redirect(ConfigurationManager.AppSettings["LoginServer"] + "myAccount/literegistration.htm?returnPage=" + context.Server.UrlEncode(returnUrl.ToString()));
            }
            else
            {
                context.Response.Redirect(ConfigurationManager.AppSettings["LoginServer"] + "myAccount/signin.htm?returnPage=" + context.Server.UrlEncode(returnUrl.ToString()));
            }
        }
    }
}
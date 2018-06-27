<!-- $Header: jtflogin.jsp 115.47 2002/02/08 16:30:51 pkm ship    $ -->
<%@ include file="jtfincl.jsp"%>
<%@ page session="false" %>
<%
  /**
   * There are different views controlled by the following parameters:
   * isLoginOnly: true or false. if true, show only login bin.
   * SecurityGroup.isHostingEnv(): true or false.
   *              if true, show orgnization field in login bin 
   *              and validate it.
   */

   String mediaPath = "/OA_MEDIA/";
   String leftCorner = mediaPath+"jtfutl02.gif";
   String rightCorner = mediaPath+"jtfutr02.gif";
   String appName = "JTF";
   boolean stateless = true;

   // whether this login is for a particular destination
   String nextPage = request.getParameter("jttNextPage");
   boolean isLoginOnly = (nextPage!=null);

   // implicitly login as a guest for regional settings
   oracle.apps.jtf.base.session.FWSession ss = null;
   try {
     ServletSessionManager.startRequest(request,response,appName,stateless);
   }catch (FrameworkException uae){
     String key = uae.getKey();
     //CredentialIncorrectException, Session Timeout, UnAuthenticatedUser, Maintenance Mode
     if(key.equals("JTF-1009") || key.equals("JTF-0177") || key.equals("JTF-0198") || key.equals("JTF-0162") )
     {
       try{
         ServletSessionManager.startStandAloneSession(appName, stateless);
       }catch(FrameworkException e){
         key = e.getKey();
         //CredentialIncorrectException 
         if(key.equals("JTF-1009")){
            WebAppsContext wctx = new WebAppsContext(java.lang.System.getProperty("JTFDBCFILE"));
            String g_userName = null;
            String g_pwd = null;
            if(wctx != null){
              String guestpwd = wctx.getEnvStore().getEnv("GUEST_USER_PWD");
              StringTokenizer st = new StringTokenizer(guestpwd,"/");
              g_userName = (String)st.nextToken();
              g_pwd = (String)st.nextToken();
		
              ServletSessionManager.startStandAloneSession(appName, stateless, g_userName, g_pwd);
            }
         }// end of if
       }// end of catch of CredentialIncorrectException..
     }else{// end of main if..
        throw uae;
     }
   }

   // get regional prompts and messages, etc
   final String[] usPrompts = {"Login", "User ID", "Password", "Go", "Register Here"};
   final String errMesg1 = "Login failed.  Please check your User ID / Password.";

   String[] prompts = usPrompts;
   String message001 = errMesg1;
   try  {
      int respid = Integer.parseInt(ServletSessionManager.getCookieValue(JTFCookie.RESP_ID));
      String langc = ServletSessionManager.getCookieValue(JTFCookie.DEFAULT_LANGUAGE); 

      //out.println("respid = "+respid);
      //out.println("lang = "+langc);

      if(langc==null) langc= "US";
      String[] pmts =
        oracle.apps.jtf.util.UIUtil.getRegionPrompts(
           "JTFLOGIN",
           respid,
           "JTF",
           langc);
      if(pmts!=null && pmts.length==prompts.length)
        prompts = pmts;
      message001 = 
        oracle.apps.jtf.util.UIUtil.getMessage(
           "JTFLOGIN001",
           errMesg1);
   }catch(Exception eee) {
     prompts = usPrompts;
   }

   // these are lines used to construct a login form
   String formHeaderLine =
     "<form name=\"login\" onsubmit=\"return validate();\" action=\"jtfavald.jsp\" method=\"POST\">";
   String formUserIDLine =
     "<input type=\"TEXT\" size=\"10\" name=\"username\" value=\""+
     (request.getParameter("username")==null?"":HtmlWriter.preformat(request.getParameter("username")))+
     "\">";
   String formPasswordLine =
     "<input type=\"PASSWORD\" size=\"10\" name=\"password\" value=\"\">";
   String formOrgLine =
     "<input type=\"TEXT\" size=\"10\" name=\"jtt_secgrpkey\" value=\"\">";
   String formSubmitLine = 
     "<input type=\"SUBMIT\" value=\""+prompts[3]+"\" name=\"enter\">";
   String formTailLine = "</form>";

   // lines for other stuff
   String registrationLine = 
     "<a href='jtfrreg1.jsp?target=jtflogin.jsp&jtt_uua=n' >"+prompts[4]+"</a>";


%>

<html>
<head>
<title>Oracle CRM</title>

<!-- Setup Style Sheet for user session --> 
<%@ include file='jtfscss.jsp'  %> 
</head>

<body bgcolor="#ffffff"  link='#663300' alink='FF6600' vlink='#996633' >
<%  
{ // legacy. someone wished to branch it
%>
<table border=0 width="100%">
  <% /* Header image */ %>
  <%@ include file="jtfloginh.jsp" %>
<%
  // print out error message if error is set
  String errorMsg = request.getParameter("error");
  if(errorMsg!=null && errorMsg.trim().length()>0) {
    out.println("<tr><td nowrap colspan=2 align=center class=errorMessage>");
    out.println(message001);
    out.println("</td></tr>");
  }

  String promptMessage = request.getParameter("jtt_mesg");
  if ( promptMessage != null && promptMessage.trim().length() > 0) {
      promptMessage = 
          oracle.apps.jtf.util.UIUtil.getMessage(promptMessage, promptMessage);
      out.println("<tr><td colspan=2 align=center class=prompt>");
      out.println(promptMessage);
      out.println("</td><tr>");
  }
%>
  <tr>
<%
  if(!isLoginOnly) {
%>
    <td valign=top>
      <!-- Login Bin Start -->
      <table width="225" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td> <!--This table creates the rounded bin header.-->
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr> 
                <td><img src="<%=leftCorner%>" width="15" height="25"></td>
                <td align="center" nowrap width="100%" class="binHeaderCell"><%=prompts[0]%></td>
                <td><img src="<%=rightCorner%>" width="15" height="25"></td>
              </tr>
            </table>
          </td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#999999"> 
            <table width="100%" border="0" cellspacing="1">
              <tr> 
                <td class="binContentCell">
                  <table border="0" cellpadding="5" cellspacing="0">
                    <%=formHeaderLine%>
                      <tr> 
                        <td align=right class="sectionHeaderBlack"><%=prompts[1]%></td>
                        <td>
                          <%=formUserIDLine%> 
                        </td>
                      </tr>
<% if(SecurityGroup.isHostingEnv()) { %>
                      <tr> 
                        <td align=right class="sectionHeaderBlack"><%=prompts[2]%></td>
                        <td nowrap>
                           <%=formPasswordLine%>
                        </td>
                      </tr>
                      <tr> 
                        <td align=right class="sectionHeaderBlack">Organization</td>
                        <td>
                           <%=formOrgLine%>
                           <%=formSubmitLine%>
                        </td>
                      </tr>
<% } else { %>
                      <tr> 
                        <td align=right class="sectionHeaderBlack"><%=prompts[2]%></td>
                        <td>
                           <%=formPasswordLine%>
                           <%=formSubmitLine%>
                        </td>
                      </tr>
<% } %>
                      <tr>
                          <td colspan='2'>
                             <%=registrationLine%>
                          </td>
                      </tr>
		      <tr>
                          <td colspan='2'>
                            <%@ include file='jtfResetPwdLink.jsp'  %>
                          </td>
                      </tr> 
                    <%=formTailLine%>
                  </table>
                </td>
              </tr>
            </table>
          </td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <!-- Login Bin End -->
      <br>
      <!-- Custom Bins -->
      <%@ include file="jtfloginf.jsp" %>
    </td>
    <td valign=top>
      <!-- News -->
      <%@ include file="jtfloginb.jsp" %>
    </td>
<%
  }else {
%>
<td nowrap colspan="2" align="center">
<br><br><br>
  <table border=0 cellspacing=0 width="60%">
    <tr><td class="homeBigHeaderCell"><%=prompts[0]%></td></tr>
    <tr><td class="binContentCell">
    <table border=0 cellspacing=5 width="100%">
      <tr><td colspan=2><p>&nbsp;</td></tr>
      <%=formHeaderLine%>
      <tr>
       <td align="right" valign="middle" class="sectionHeaderBlack">
          <%=prompts[1]%></td>
       <td><%=formUserIDLine%></td>
      </tr>
<% if(SecurityGroup.isHostingEnv()) { %>
      <tr>
       <td align="right" valign="middle" class="sectionHeaderBlack">
          <%=prompts[2]%></td>
       <td><%=formPasswordLine%></td> 
      </tr>
      <tr>
       <td align="right" valign="middle" class="sectionHeaderBlack">
          Organization</td>
       <td>
         <%=formOrgLine%>
         <%=formSubmitLine%>
       </td> 
      </tr>
<% } else { %>
      <tr>
       <td align="right" valign="middle" class="sectionHeaderBlack">
          <%=prompts[2]%></td>
       <td>
          <%=formPasswordLine%>
          <%=formSubmitLine%>
       </td> 
      </tr>
<% } %>
<%
      if(nextPage != null) out.println("<input type='hidden' name='jttNextPage' value='"+ nextPage + "'>");
%>
      <%=formTailLine%>
      <tr><td colspan=2><p>&nbsp;</td></tr>
    </table> 
    </td></tr>
  </table>
</td>
<%
  } // end of if(!isLoginOnly)
%>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

</table>


<% } %>

<SCRIPT LANGUAGE="JavaScript">

<% if(SecurityGroup.isHostingEnv()) { %>

 // Enter key on last field results in submission
 document.login.jtt_secgrpkey.onkeypress = keyhandler;
 
 function validate() {
   if(document.login.jtt_secgrpkey.value == null || document.login.jtt_secgrpkey.value =="" || document.login.jtt_secgrpkey.value.toUpperCase() == 'STANDARD')
   {
     alert('Enter a valid Organization');
     return false;
   } else {
     return true;
   }
 }

<% }else{ %>

 // Enter key on last field results in submission
 document.login.password.onkeypress = keyhandler;

 function validate() {
   return true;
 }

<% } %>

 function keyhandler(e) {
   if(navigator.appName == 'Netscape'){
      if (document.layers)
          Key = e.which;
      else
          Key = window.event.keyCode;
      if (Key == 13) {
        if(validate()) document.login.submit();
      }
   }
 }

</SCRIPT>
</body>
<!-- End the session. --> 
<% ServletSessionManager.endStandAloneSession(); %>
<HEAD>
<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
</HEAD>
</html>

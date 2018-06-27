<%@ Language="VBScript" CODEPAGE="65001" %>

<%
' --------------------------------------------------------------------------------------
' Author:     D. Whitlock
' Date:       2/1/2000
' --------------------------------------------------------------------------------------

response.buffer = true

%>
<!--#include virtual="/include/functions_string.asp"-->
<!--#include virtual="/include/functions_translate.asp"-->
<!--#include virtual="/include/functions_file.asp"-->
<!--#include virtual="/SW-Common/Preferred_Language.asp"-->
<!--#include virtual="/connections/connection_SiteWide.asp"-->
<%

' --------------------------------------------------------------------------------------
' Connect to SiteWide DB
' --------------------------------------------------------------------------------------

Call Connect_SiteWide

' --------------------------------------------------------------------------------------
' Declarations
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Security_Module.asp" -->
<%

Dim BackURL
Dim LimitView
Dim ErrorString

BackURL = Session("BackURL")    

' --------------------------------------------------------------------------------------
' Determine Login Credintials and Site Code and Description based on Site_ID Number 
' --------------------------------------------------------------------------------------

%>
<!--#include virtual="/SW-Common/SW-Site_Information.asp"-->
<%

Dim Top_Navigation        ' True / False
Dim Side_Navigation       ' True / False
Dim Screen_Title          ' Window Title
Dim Bar_Title             ' Black Bar Title

Screen_Title    = Translate(Site_Description,Alt_Language,conn) & " - " & Translate("Service Documents - Help",Alt_Language,conn)
Bar_Title       = Translate(Site_Description,Login_Language,conn) & "<BR><FONT CLASS=SmallBoldGold>" & Translate("Service Documents - Help",Login_Language,conn) & "</FONT>" 
Top_Navigation  = False
Side_Navigation = True
Content_Width   = 95  ' Percent

%>
<!--#include virtual="/SW-Common/SW-Header.asp"-->
<!--#include virtual="/SW-Common/SW-Common-Navigation.asp"-->
<%

response.write "<FONT CLASS=Heading3>" & Translate("Service Documents",Login_Language,conn) & "</FONT>"
response.write "<BR><BR>"

response.write "<FONT CLASS=Medium>"

' --------------------------------------------------------------------------------------
' Main
' --------------------------------------------------------------------------------------
%>
The Fluke Service Documentation Database and Fluke Service Support Database, maintained by CSS Service Support Engineering, are resource listings of product and service related information grouped by model number.  The Fluke Service index contains a description of Product Change Notices, Service Alerts, Service Bulletins, etc.  The Fluke Service Support Index contains warranty, end of support, engineer information.  The following sections describe the columns and key information that comprise these indexes:
<P> 
<HR NOSHADE>
<STRONG>MODEL NUMBER</STRONG>
<P>
The Fluke Service Index is sorted by product Model Number, then by Document Number.  Typically, a product will only have an entry 
into the Service Index when service related information has been written relating to the Model Number
<P>
In some instances, a document may appear in the index under several model numbers.  For example, PCN 742 appears under model numbers 5100B, 
5101B and 5102B since this PCN is applicable to all three model numbers.
<P>
In the case of a special version of a model,  the index listing under that special model number will contain only PCNs and Service Alerts that apply to 
that special.  PCNs and Service Alerts that apply to the basic product, will not be repeated in this section.  For example, PCN 651 appears under the 
model number 1720A/AP since the PCN describes a unique change to the /AP special.  Therefore, the service technician would look under both the 
1720A and the 1720A/AP to find all applicable PCNs and Service Alerts.
<P>
<HR NOSHADE>
<STRONG>ASSEMBLY</STRONG>
<P>
If applicable, the effected assembly is noted for ease of reference.  If a publication does not directly address a specific assembly, the document type 
description is supplied.  This allows the user to quickly scan the index for general information on a specific model number.
<HR NOSHADE>
<STRONG>BOARD REVISION</STRONG>
<P>
The board revision column indicates the revision range of the board or assembly affected.  A special case revision is noted as “VER” for software or 
firmware updates.  The following is a key to board revision number ranges used within the index:
<P>
<DIV ALIGN=CENTER>
<TABLE BORDER NOSHADE>
<TR>
<TD CLASS=Medium>Definition</TD>
<TD CLASS=Medium>Syntax</TD>
</TR>

<TR>
<TD CLASS=Medium>Revision Range</TD>
<TD<FONT size=2 FACE="ARIAL, Verdana, Helvetica">>A--&gtC</TD>
</TR>

<TR>
<TD CLASS=Medium>Revision Range Combination</TD>
<TD CLASS=Medium>A--&gtC,E,G--&gtJ</TD>
</TR>

<TR>
<TD CLASS=Medium>Less Than</TD>
<TD CLASS=Medium>&lt   B</TD>
</TR>

<TR>
<TD CLASS=Medium>Less Than or Equal To</TD>
<TD CLASS=Medium>&lt= B</TD>
</TR>

<TR>
<TD CLASS=Medium>Greater Than</TD>
<TD CLASS=Medium>&gt   B</TD>
</TR>

<TR>
<TD CLASS=Medium>Greater Than or Equal To</TD>
<TD CLASS=Medium>&gt=  B</TD>
</TR>

<TR>
<TD CLASS=Medium>All Revisions</TD>
<TD CLASS=Medium>ALL</TD>
</TR>

<TR>
<TD CLASS=Medium>See Document for Range</TD>
<TD CLASS=Medium>SEE PCN</TD>
</TR>

<TR>
<TD CLASS=Medium>Software / Firmware Version</TD>
<TD CLASS=Medium>VER XXX</TD>
</TR>

</TABLE>
</DIV>

<P>
The board revision column also indicates publications which are written in languages other than English.  The following table notes the Language Code 
and Language.
<P>

<DIV ALIGN=CENTER>
<TABLE BORDER NOSHADE>
<TR>
<TD CLASS=Medium>Definition</TD>
<TD CLASS=Medium>Syntax</TD>
</TR>

<TR>
<TD CLASS=Medium>Chinese</TD>
<TD CLASS=Medium>[CH]</TD>
</TR>

<TR>
<TD CLASS=Medium>Danish</TD>
<TD CLASS=Medium>[DA]</TD>
</TR>

<TR>
<TD CLASS=Medium>Dutch</TD>
<TD CLASS=Medium>[DU]</TD>
</TR>

<TR>
<TD CLASS=Medium>English</TD>
<TD CLASS=Medium>[EN]</TD>
</TR>

<TR>
<TD CLASS=Medium>Finnish</TD>
<TD CLASS=Medium>[FI]</TD>
</TR>

<TR>
<TD CLASS=Medium>French</TD>
<TD CLASS=Medium>[FR]</TD>
</TR>

<TR>
<TD CLASS=Medium>German</TD>
<TD CLASS=Medium>[GE]</TD>
</TR>

<TR>
<TD CLASS=Medium>Italian</TD>
<TD CLASS=Medium>[IT]</TD>
</TR>

<TR>
<TD CLASS=Medium>Japanese</TD>
<TD CLASS=Medium>[JA]</TD>
</TR>

<TR>
<TD CLASS=Medium>Korean</TD>
<TD CLASS=Medium>[KO]</TD>
</TR>

<TR>
<TD CLASS=Medium>Multiple Language</TD>
<TD CLASS=Medium>[ML]</TD>
</TR>

<TR>
<TD CLASS=Medium>Noregian</TD>
<TD CLASS=Medium>[NO]</TD>
</TR>

<TR>
<TD CLASS=Medium>Spanish</TD>
<TD CLASS=Medium>[SP]</TD>
</TR>

<TR>
<TD CLASS=Medium>Swedish</TD>
<TD CLASS=Medium>[SW]</TD>
</TR>

</TABLE>

</DIV>

<HR NOSHADE>
<STRONG>SERIAL NUMBER</STRONG>
<P>
The Service Index lists the serial number or serial number range where a document incorporated the change.  Please note as an example, that the serial 
number is often approximated, since many of the changes are phased into production after a document was written.
<P>

<DIV ALIGN=CENTER>
<TABLE BORDER NOSHADE>
<TR>
<TD CLASS=Medium>Definition</TD>
<TD CLASS=Medium>Syntax</TD>
</TR>

<TR>
<TD CLASS=Medium>Range</TD>
<TD CLASS=Medium>00000000-00000000</TD>
</TR>

<TR>
<TD CLASS=Medium>Less Than</TD>
<TD CLASS=Medium>&lt   00000000</TD>
</TR>

<TR>
<TD CLASS=Medium>Less Than or Equal To</TD>
<TD CLASS=Medium>&lt= 00000000</TD>
</TR>

<TR>
<TD CLASS=Medium>Greater Than</TD>
<TD CLASS=Medium>&gt   00000000</TD>
</TR>

<TR>
<TD CLASS=Medium>Greater Than or Equal To</TD>
<TD CLASS=Medium>&gt=  00000000</TD>
</TR>

<TR>
<TD CLASS=Medium>Not Applicable</TD>
<TD CLASS=Medium>NA</TD>
</TR>

<TR>
<TD CLASS=Medium>All Serial Numbers</TD>
<TD CLASS=Medium>ALL</TD>
</TR>

<TR>
<TD CLASS=Medium>See Document for Range</TD>
<TD CLASS=Medium>SEE PCN</TD>
</TR>

</TABLE>

</DIV>

<HR NOSHADE>
<STRONG>DOC NUMBER</STRONG>
<P>
The document number column indicates the type of document or type of service information listed and the document number.  Documents that have a 
Part Number / Order Code can be ordered directly from the Service Parts/Exchange Centers.  Documents that do not have a Part Number / Order Code 
can be obtained by contacting the CSS Service Support Engineering -- Everett for VDT or STD products, or CSS Service Support Engineering -- Almelo 
for DTD products.  Document Types listed in the Service Index are as follows:

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">PCN</FONT> <EM>PRODUCT CHANGE NOTICE</EM></STRONG>
<P>
Product Change Notices describe product modifications which must be installed for enhancement, reliability improvement or for safety 
reasons.  PCNs are written and distributed by CSS Service Support Engineering.  PCNs are grouped into five classes:
<P>
<B>Class 0</B> <EM>“Mandatory Change”</EM>  Normally used for safety related modifications, though may be used for any modification deemed 
mandatory.  Installed in all instruments received by Service, regardless of the age of the instrument, and is effective until the 
PCN is canceled.
<P>
<B>Class 1</B> <EM>“Reliability Improvement”</EM>  Modification installed in all instruments received by Service, regardless if symptom or 
condition is observed, for a period of one year beyond the normal warranty period.
<P>
After one year beyond the normal warranty period, Class 1 PCNs are installed only if the PCN will repair the problem the 
instrument was sent in for.
<P>
<B>Class 2</B> <EM>“Reliability Improvement”</EM> Modifications installed in all instruments received by Service for a period of one year beyond 
the normal warranty period, only if the instrument will display the symptom or condition described in the PCN.
<P>
After one year beyond the normal warranty period, Class 2 PCNs are installed only if the PCN will repair the problem the 
instrument was sent in for.
<P>
<B>Class 3</B> <EM>“Parts Substitutions”</EM> Component substitution or minor specification/reliability improvement.  Installed or applied only if 
the PCN  will repair the problem the instrument was sent in for.
<P>
<B>Class 4</B> <EM>“Informational Only”</EM> Notifications similar to “change errata sheets” for manuals.  There are no installation criteria.
<P>
Reference:	SOP 111.46, Warranty Policy Administration

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">SA </FONT><EM>SERVICE ALERT</EM></STRONG>
<P>
Service Alerts are confidential service information written for internal company use only, primarily to note changes to service policy or 
procedures.  Service Alerts are written and distributed by CSS Service Support Engineering.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">Change Notices</FONT><EM>, Change Information and Information Sheets</EM></STRONG>
<P>These documentsdescribe product modifications for enhancement, reliability improvement or 
for safety reasons.  These publications are written and distributed <U>only</U> by <B>CSS Service Support Engineering -- Almelo</B>.  Some of these 
publications are supplied with class codes.  These class codes should be interpreted the same as PCN class codes as described in the 
preceding pages under section, DOCUMENT NUMBER -- PCN.
<P>

<DIV ALIGN=CENTER>
<TABLE BORDER NOSHADE>
<TR>
<TD CLASS=Medium>Document Type</TD>
<TD CLASS=Medium>Document Description</TD>
<TD CLASS=Medium>Division</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">OSC</font></TD>
<TD CLASS=Medium>CHANGE NOTICE--OSCILLOSCOPES</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">SPC</font></TD>
<TD CLASS=Medium>CHANGE NOTICE--PULSE GENERATORS AND COUNTERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">SME</font></TD>
<TD CLASS=Medium>CHANGE NOTICE--VOLTMETERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">SSY</font></TD>
<TD CLASS=Medium>CHANGE NOTICE--SYSTEMS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">SRE</font></TD>
<TD CLASS=Medium>CHANGE NOTICE--RECORDERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">CIH</font></TD>
<TD CLASS=Medium>CHANGE INFORMATION HARDWARE--LOGIC ANALYZERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">CIS</font></TD>
<TD CLASS=Medium>CHANGE INFORMATION SOFTWARE--LOGIC ANALYZERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>

<TR>
<TD><FONT COLOR="#FF0000"><FONT size=2 FACE="ARIAL, Verdana, Helvetica">DTE</font></TD>
<TD CLASS=Medium>INFORMATION SHEET--LOGIC ANALYZERS</TD>
<TD CLASS=Medium>(DTD Products Only)</TD>
</TR>
</TABLE>
</DIV>

<HR NOSHADE>

<STRONG><FONT COLOR="#FF0000">SBU</FONT> <EM>SERVICE BULLETIN</EM> </STRONG>(DTD Products Only)
<P>
Service Bulletins are confidential service information on DTD products written for internal company use only.  Service Bulletins are written 
and distributed by CSS Service Support Engineering -- Almelo.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">SD</FONT> <EM>SERVICE DOCUMENTATION</EM></STRONG>
<P>
Detailed product repair and maintenance information.  The current revision of the service documentation is listed with part number / order 
code.  Service documentation (manuals) are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">SUP</FONT> <EM>SERVICE SUPPLEMENT -- GENERAL</EM></STRONG>
<P>
This publication contains supplemental repair and maintenance information not found in the Service Manual.  SUPs are written and 
distributed <U>only</U> by <B>Service Support Engineering -- Everett</B>.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">SSU</FONT> <EM>SERVICE SUPPLEMENT -- DTD</EM></STRONG> (DTD Products Only)
<P>	
This publication contains supplemental repair and maintenance information not found in the Service Manual. The current revision of the 
service supplement is listed with part number / order code.  SSUs are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">UD</FONT> <EM>USER DOCUMENTATION</EM></STRONG>
<P>
Detailed product user information. The current revision of the user documentation is listed with part number / order code.  Manuals are 
ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">USU</FONT> <EM>USER SUPPLEMENT -- DTD</EM></STRONG> (DTD Products Only)
<P>	
This publication contains supplemental information not found in the user documentation. The current revision of the user supplement is 
listed with part number / order code.  USUs are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">MSU</FONT> <EM>MANUAL SUPPLEMENT</EM></STRONG>
<P>
This publication contains supplemental information not found in the user or service documentation. The current revision of the manual 
supplement is listed with part number / order code.  MSUs are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">ESU</FONT> <EM>SUPPLEMENT -- OSCILLOSCOPE MANUALS</EM></STRONG> (DTD Products Only)
<P>	
This publication contains supplemental information not found in the user or service documentation. The current revision of the user 
supplement is listed with part number / order code.  ESUs are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">IM</FONT> <EM>INSTRUCTION MANUAL</EM></STRONG> (DTD Products Only)
<P>
Detailed product user information.  Some publications also include detailed product repair and maintenance information. The current 
revision of the user documentation is listed with part number / order code.  Instruction Manuals are ordered through the Service 
Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">KIT</FONT> <EM>PARTS KIT</EM></STRONG>
<P>
Parts Kits include modification kits, parts replacement kits, upgrade kits, and option kits.  The kit part number is shown in part number / 
order code column.  Kits are ordered through the Service Parts/Exchange Centers.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">CLASS</FONT> <EM>PCN CLASS CODE</EM></STRONG>
<P>
Class Codes apply only to Product Change Notices (PCN) and DTD Change Notices.  The five (5) PCN classes described in the preceding pages under 
section, DOCUMENT NUMBER -- PCN.

<HR NOSHADE>
<STRONG><FONT COLOR="#FF0000">DATE</FONT> <EM>DOCUMENT DATE</EM></STRONG>
<P>
The date column shows the original date that the document was created.
<HR NOSHADE>

<STRONG><FONT COLOR="#FF0000">REV</FONT> <EM>DOCUMENT REVISION DATE</EM></STRONG>
<P>
This column lists the current revision date of the document or manual.

<HR NOSHADE>
<STRONG>PN / ORDER CODE</STRONG>
<P>
This column lists either a 6 digit part number or 12 digit order code number, if the document or publication can be ordered through the Service 
Parts/Exchange Centers.  The Service Index only list publications written in English or a Multi-Language publication that includes English and 
publications (manuals, etc.) for products for which a service document initiated an entry into the Service Index.
<P>
User or Service publications that are indicated as “In-Active” are kept in the Service Index for archive purposes, however copies of these documents are 
no longer supported and may not be available through the Service Parts/Exchange Centers.
<P>
To see other language versions of a publication available for VTD and STD products, reference the “List of Available User Documentation”, provided 
by Publication Services -- Everett, c/o Maria Beers, MS 232-E.
<P>
To see other language versions of a publication available for DTD products, reference the “DTD Publication Index”, provided by Customer Support Services -- Almelo, c/o Thom van der Vat.

<HR NOSHADE>
<STRONG>SUBJECT OR ADDITIONAL INFORMATION</STRONG>
<P>
The subject or additional information column is a brief description of the subject matter of the publication.

<HR NOSHADE>
<STRONG>PRODUCTION ENDED</STRONG>
<P>
If a product has gone out-of-production, the effective Production Ended date is listed.  This date is useful in the determination of the products “Long 
Term Support Period”.   The last year of the “Long Term Support Period” [year] is noted in brackets, following the Production Ended date.

<HR NOSHADE>
<STRONG>WARRANTY</STRONG>
<P>
The product’s normal warranty period in years.
</P>

<BR><BR>

<!--#include virtual="/SW-Common/SW-Footer.asp"-->

<%
Call Disconnect_SiteWide
%>
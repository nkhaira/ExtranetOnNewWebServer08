/*
   Monday, January 15, 2007 1:48:43 PM
   User: Kelly Whitlock
   Server: FLKTST18 and FLKPRD18
   Database: Fluke_SiteWide
   Application: MS SQLEM - Data Tools
*/

USE Fluke_SiteWide

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
ALTER TABLE dbo.Calendar
	DROP CONSTRAINT DF_Calendar_PID
GO
ALTER TABLE dbo.Calendar
	DROP CONSTRAINT DF_Calendar_Cost_Center
GO
ALTER TABLE dbo.Calendar
	DROP CONSTRAINT DF_Calendar_File_Pages
GO
ALTER TABLE dbo.Calendar
	DROP CONSTRAINT DF_Calendar_Secure_Stream
GO
ALTER TABLE dbo.Calendar
	DROP CONSTRAINT DF_Calendar_Subscription_Early
GO
CREATE TABLE dbo.Tmp_Calendar
	(
	ID int NOT NULL IDENTITY (1, 1),
	Site_ID int NULL,
	Code int NULL,
	Category_ID int NULL,
	Sub_Category nvarchar(255) NULL,
	Content_Group smallint NULL,
	Content_Group_Name nvarchar(100) NULL,
	Product nvarchar(255) NULL,
	Title nvarchar(255) NULL,
	Description nvarchar(2000) NULL,
	Instructions nvarchar(500) NULL,
	Splash_Header nvarchar(1000) NULL,
	Splash_Footer nvarchar(500) NULL,
	Item_Number varchar(20) NULL,
	Item_Number_Show smallint NULL,
	Item_Number_2 varchar(20) NULL,
	PID numeric(18, 0) NULL,
	Revision_Code varchar(4) NULL,
	Cost_Center smallint NULL,
	Location nvarchar(255) NULL,
	LDays int NULL,
	LDate datetime NULL,
	BDate datetime NULL,
	EDate datetime NULL,
	VDate datetime NULL,
	XDate datetime NULL,
	XDays int NULL,
	PDate datetime NULL,
	UDate datetime NULL,
	PEDate datetime NULL,
	Confidential smallint NULL,
	Link varchar(255) NULL,
	Link_PopUp_Disabled smallint NULL,
	Include varchar(255) NULL,
	Include_Size int NULL,
	File_Name varchar(255) NULL,
	File_Size int NULL,
	File_Page_Count int NULL,
	Archive_Name varchar(255) NULL,
	Archive_Size int NULL,
	File_Name_POD varchar(255) NULL,
	File_Size_POD int NULL,
	Archive_Name_POD varchar(255) NULL,
	Archive_Size_POD int NULL,
	Thumbnail varchar(255) NULL,
	Thumbnail_Size int NULL,
	Thumbnail_Request smallint NULL,
	Secure_Stream smallint NULL,
	Image_Locator varchar(255) NULL,
	Forum_ID int NULL,
	Forum_Moderated smallint NULL,
	Forum_Moderator_ID int NULL,
	Groups varchar(50) NULL,
	SubGroups varchar(2048) NULL,
	Subscription smallint NULL,
	Subscription_Early smallint NULL,
	Headline_View smallint NULL,
	[Language] varchar(10) NULL,
	Country varchar(1000) NULL,
	Status smallint NULL,
	Status_Override smallint NULL,
	Status_Comment varchar(255) NULL,
	Clone int NULL,
	Locked smallint NULL,
	Submitted_By int NULL,
	Approved_By int NULL,
	Review_By int NULL,
	Review_By_Group int NULL,
	Campaign int NULL
	)  ON [PRIMARY]
GO
DECLARE @v sql_variant 
SET @v = N'EEF Preview Date'
EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'Tmp_Calendar', N'column', N'VDate'
GO
DECLARE @v sql_variant 
SET @v = N'Direct File Access via URL or Secure Stream'
EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'Tmp_Calendar', N'column', N'Secure_Stream'
GO
DECLARE @v sql_variant 
SET @v = N'If True, Subscription is sent mid-day PST'
EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'Tmp_Calendar', N'column', N'Subscription_Early'
GO
DECLARE @v sql_variant 
SET @v = N'Overrides Oracle Status Intervention until Oracle Item Status is LIVE'
EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'Tmp_Calendar', N'column', N'Status_Override'
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_PID DEFAULT (0) FOR PID
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_Cost_Center DEFAULT (0) FOR Cost_Center
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_File_Pages DEFAULT (0) FOR File_Page_Count
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_Secure_Stream DEFAULT (0) FOR Secure_Stream
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_Subscription_Early DEFAULT (0) FOR Subscription_Early
GO
ALTER TABLE dbo.Tmp_Calendar ADD CONSTRAINT
	DF_Calendar_Status_Override DEFAULT 0 FOR Status_Override
GO
SET IDENTITY_INSERT dbo.Tmp_Calendar ON
GO
IF EXISTS(SELECT * FROM dbo.Calendar)
	 EXEC('INSERT INTO dbo.Tmp_Calendar (ID, Site_ID, Code, Category_ID, Sub_Category, Content_Group, Content_Group_Name, Product, Title, Description, Instructions, Splash_Header, Splash_Footer, Item_Number, Item_Number_Show, Item_Number_2, PID, Revision_Code, Cost_Center, Location, LDays, LDate, BDate, EDate, VDate, XDate, XDays, PDate, UDate, PEDate, Confidential, Link, Link_PopUp_Disabled, Include, Include_Size, File_Name, File_Size, File_Page_Count, Archive_Name, Archive_Size, File_Name_POD, File_Size_POD, Archive_Name_POD, Archive_Size_POD, Thumbnail, Thumbnail_Size, Thumbnail_Request, Secure_Stream, Image_Locator, Forum_ID, Forum_Moderated, Forum_Moderator_ID, Groups, SubGroups, Subscription, Subscription_Early, Headline_View, [Language], Country, Status, Status_Comment, Clone, Locked, Submitted_By, Approved_By, Review_By, Review_By_Group, Campaign)
		SELECT ID, Site_ID, Code, Category_ID, Sub_Category, Content_Group, Content_Group_Name, Product, Title, Description, Instructions, Splash_Header, Splash_Footer, Item_Number, Item_Number_Show, Item_Number_2, PID, Revision_Code, Cost_Center, Location, LDays, LDate, BDate, EDate, VDate, XDate, XDays, PDate, UDate, PEDate, Confidential, Link, Link_PopUp_Disabled, Include, Include_Size, File_Name, File_Size, File_Page_Count, Archive_Name, Archive_Size, File_Name_POD, File_Size_POD, Archive_Name_POD, Archive_Size_POD, Thumbnail, Thumbnail_Size, Thumbnail_Request, Secure_Stream, Image_Locator, Forum_ID, Forum_Moderated, Forum_Moderator_ID, Groups, SubGroups, Subscription, Subscription_Early, Headline_View, [Language], Country, Status, Status_Comment, Clone, Locked, Submitted_By, Approved_By, Review_By, Review_By_Group, Campaign FROM dbo.Calendar TABLOCKX')
GO
SET IDENTITY_INSERT dbo.Tmp_Calendar OFF
GO
DROP TABLE dbo.Calendar
GO
EXECUTE sp_rename N'dbo.Tmp_Calendar', N'Calendar', 'OBJECT'
GO
COMMIT

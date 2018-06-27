
CREATE PROCEDURE [dbo].[Lit_Status_Build] AS

Declare @my_error varchar(128)
	,@start_trans integer
	,@my_rows integer

set @start_trans = 0

CREATE TABLE #tmp_lit (
	[ID] [int] NOT NULL ,
	[Item_Number] [varchar] (50) NOT NULL ,
	[Status] [int] NOT NULL ,
	[POD_file] [varchar] (255) NULL ,
	[POD_size] [int] NULL ,
	[Eff_size] [int] NULL ,
	[udate] [datetime] not null)

create TABLE #tmp_lit2 (
	[Item_Number] [varchar] (50) NOT NULL ,
	[Lcnt] [int] NOT NULL)

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Error creating table'
	GOTO error_rtn
END

insert into #tmp_lit ([id],item_number,status,pod_file,pod_size,eff_size,udate)
	select [id],item_number,status,upper(right(file_name_pod,len(file_name_pod) - CHARINDEX('/',file_name_pod))),file_size_pod,file_size,udate
	FROM  dbo.Calendar
	WHERE (Item_Number is not null)
	AND (Status = 1)
	AND (SubGroups LIKE '%view%') 
	AND (BDate < GETDATE()) 
	AND ((xdays = 0) OR (DATEADD(d, XDays, EDate) > GETDATE()))
	AND ([file_name] like '%.pdf')
	AND (site_id in (SELECT [ID] FROM Site WHERE (Enabled = - 1)))

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'First insert into tmp_lit failed'
	GOTO error_rtn
END

insert into #tmp_lit ([id],item_number,status,pod_file,pod_size,eff_size,udate)
	select [id],item_number,0 as status,upper(right(file_name_pod,len(file_name_pod) - CHARINDEX('/',file_name_pod))),file_size_pod,file_size,udate
	FROM  dbo.Calendar
	WHERE (Item_Number is not null)
	AND (SubGroups LIKE '%view%') 
	AND ([file_name] like '%.pdf')
	AND (site_id in (SELECT [ID] FROM Site WHERE (Enabled = - 1)))
	AND item_number not in (select item_number from #tmp_lit)

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Second insert into tmp_lit failed'
	GOTO error_rtn
END

insert #tmp_lit2
select item_number,count(*)
from #tmp_lit
group by item_number

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Insert into tmp_lit2 failed'
	GOTO error_rtn
END

delete #tmp_lit2 where Lcnt = 1

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Delete from tmp_lit2 failed'
	GOTO error_rtn
END

declare @maxudate datetime
	,@lit_item varchar(32)
	,@tmpcnt int
	,@rcnt int

while (1=1)
  BEGIN
	set rowcount 1
	select @lit_item = Item_Number
		, @tmpcnt = Lcnt
	from #tmp_lit2

	if (@@ROWCOUNT = 0)
	  BEGIN
		set rowcount 0
		break
	  END
	delete #tmp_lit2

	set rowcount 0

	delete  #tmp_lit
	where item_number = @lit_item and pod_file is null
	set @rcnt = @@ROWCOUNT

	if (@tmpcnt - @rcnt <> 1)
	  BEGIN
		select @maxudate = max(udate) from #tmp_lit where item_number = @lit_item
	
		delete #tmp_lit
		where item_number = @lit_item and udate <> @maxudate
	  END

	IF (@@ERROR <> 0)
	BEGIN
		set @my_error = 'Error deleting from tmp_lit on date'
		GOTO error_rtn
	END
  END

begin transaction
set @start_trans = 1

delete lit_status

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Error deleting old data'
	GOTO error_rtn
END

insert into lit_status ([id],item_number,status,pod_file,pod_size,eff_size)
select [id],item_number,status,pod_file,pod_size,eff_size from #tmp_lit

set @my_rows = @@ROWCOUNT

IF (@@ERROR <> 0)
BEGIN
	set @my_error = 'Error in big copy of new data'
	GOTO error_rtn
END
commit

set @my_error = convert(varchar(64),@my_rows) + ' rows inserted'
select @my_error
return

error_rtn:
	if (@start_trans = 1) rollback tran
	RAISERROR(@my_error,16,1)
GO

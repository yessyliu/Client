create table Client(
	ClientID varchar(10) primary key
	,Name varchar(100)
	,Gender varchar(1)
	,DOB datetime
	,CreatedBy varchar(20)
	,CreatedDate datetime
	,UpdatedBy varchar(20)
	,UpdatedDate datetime
)

create function dbo.fnGenClientID()
returns varchar(10)
as
begin
	declare @id int, @Hasil varchar(10)
	select @id = isnull(max(convert(int,right(ClientID,4))),0) from Client where ClientID like ''+right(convert(varchar,getdate(),112),6)+'%'
	select @Hasil=right(convert(varchar,getdate(),112),6)+Right('0000' + convert(varchar,@id+1), 4) 
	return @Hasil
end
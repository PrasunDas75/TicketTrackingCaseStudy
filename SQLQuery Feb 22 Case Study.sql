--create database db_ticketTracking
--use db_ticketTracking
-------------------------------------
create table EMPLOYEE(
	EID nvarchar(20) primary key,
	Employee_Name nvarchar(50) not null,
	Hire_Date date not null,
	Dept nvarchar(20)
)
-------------------------------------
create table EmployeeAuthentication(
	AuthId int identity(1,1) primary key,
	EID nvarchar(20) foreign key references EMPLOYEE(EID) unique,
	[User_ID] nvarchar(50),
	[Password] nvarchar(60)
)
-------------------------------------
Create table TICKETS(
	Ticket_Id int identity(1,1) primary key,
	Logged_By nvarchar(20) foreign key references EMPLOYEE(EID),
	Raised_Date date,
	Severity varchar(15),
	Ticket_Desc varchar(100),
	Resolved_By nvarchar(20),
	Resolution varchar(100),
	Resolved_Date date,
	[Status] varchar(15)
)

ALTER TABLE TICKETS
ALTER COLUMN Resolved_Date datetime;
--------------------------------------
select * from employee
select * from employeeAuthentication
select * from tickets
SET DATEFORMAT DMY
insert into TICKETS (LOGGED_BY, RAISED_Date, Severity, TICKET_Desc, Status) values('E100100','23-02-2022' ,'fgh','hdfhdfhdfgh', 'OPEN')
insert into TICKETS (LOGGED_BY, RAISED_Date, Severity, TICKET_Desc, Status) values('E100102','23-02-2022 01:10' ,'Normal','hdfhdfhdfgsdfh', 'OPEN')

SET DATEFORMAT DMY
insert into TICKETS (LOGGED_BY, RAISED_Date, Severity, TICKET_Desc, Status) values('E100101','23-02-2022','ghfg','dhdfhdfhd', 'OPEN')

insert into TICKETS  values('E100103','23-02-2022 00:10','Major','dhdfhdfhd','M100103','solved','23-02-2022 16:50', 'CLOSED')

truncate table tickets

--select Format(GETDATE(), 'mm-dd-yyyy hh:mm')
-----------------------adding values--------------------------------------------
insert into EMPLOYEE values   
	('E100100','Venkat','2004-1-10','MGM')
,	('E100101','Krishna','2004-1-10','MGM')
,   ('E100102','Chandrashekhar','2005-3-11','DEV')
,   ('E100103','Saheer Ali Khan','2008-10-13','DEV')
,   ('E100104','Shashikanth','2007-2-17','DEV')
,   ('M100103','Avinash','2007-3-10','DEVOPS')
,   ('M100105','Ashok','2008-6-18','DEVOPS')
---------------------------------------------------
insert into EmployeeAuthentication (EID, [User_ID], [Password]) values   
	('E100100','Venkat','Venkat@123')
,	('E100101','Krishna','Krishna@123')
,   ('E100102','Chandrashekhar','Chandrashekhar@123')
,   ('E100103','Saheer Ali Khan','Saheer@123')
,   ('E100104','Shashikanth','Shashi@123')
,   ('M100103','Avinash','Avinash@123')
,   ('M100105','Ashok','Ashok@123')

--truncate table EmployeeAuthentication
-------------------------------------------------------

select ea.AuthId from EmployeeAuthentication as ea
inner join EMPLOYEE as e
on e.EID = ea.EID 
where ea.EID = 'E100100' and ea.[Password]='Venkat@123' and e.Dept = 'MGM'

-------------------------------------------------------

--alter function ViewTickets() returns table
--as
--	return
--	select e.Employee_Name as 'Employee Name',t.Ticket_Id as 'Ticket',t.Severity as 'Severity',DATEDIFF(HOUR,t.Raised_Date,t.Resolved_Date) as 'Turnaround Time',t.Ticket_Desc as 'Description',t.Resolved_By as 'Resolved By' from EMPLOYEE as e
--	inner join TICKETS as t
--	on t.Logged_By = e.EID
--	where t.Status = 'CLOSED'

--select * from ViewTickets()

alter procedure ViewTicket
as
begin
	declare @ticket table(
		[Employee Name] nvarchar(50),
		[Ticket] int,
		[Severity] varchar(15),
		[Turnaround Time] float,
		[Description] varchar(100),
		[Resolved By] nvarchar(20)
	)

	insert @ticket
	select e.Employee_Name,t.Ticket_Id,t.Severity, CAST(DATEDIFF(MINUTE ,t.Raised_Date , t.Resolved_Date )/60.0 AS NUMERIC(10, 2)),t.Ticket_Desc,t.Resolved_By from EMPLOYEE as e
	inner join TICKETS as t
	on t.Logged_By = e.EID
	where t.Status = 'CLOSED'

	--select e.Employee_Name as 'Employee Name',t.Ticket_Id as 'Ticket',t.Severity as 'Severity', CAST(DATEDIFF(MINUTE ,t.Raised_Date , t.Resolved_Date )/60.0 AS NUMERIC(10, 2)) as 'Turnaround Time',t.Ticket_Desc as 'Description',t.Resolved_By as 'Resolved By' from EMPLOYEE as e
	--inner join TICKETS as t
	--on t.Logged_By = e.EID
	--where t.Status = 'CLOSED'

	select t.*,e.Employee_Name  from @ticket as t
	inner join EMPLOYEE as e
	on e.EID = t.[Resolved By] 
end
--select DATEDIFF(HOUR,'22-02-2022','23-02-2022')

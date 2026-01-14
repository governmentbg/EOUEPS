BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '6';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-01-26_Version6';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

CREATE TABLE [dbo].[Departments](
    [DepartmentId]      INT             NOT NULL IDENTITY,
    [Alias]             NVARCHAR(100)   NOT NULL,
    [Name]              NVARCHAR(200)   NOT NULL,
 CONSTRAINT [PK_Departments] PRIMARY KEY ([DepartmentId])
);


SET IDENTITY_INSERT [dbo].[Departments] ON 

EXEC('
INSERT [dbo].[Departments] ([DepartmentId], [Alias], [Name])
VALUES (1, N''MinistryOfJustice'', N''Министерство на правосъдието'')');

EXEC('
INSERT [dbo].[Departments] ([DepartmentId], [Alias], [Name])
VALUES (2, N''NACID'', N''Националният център за информация и документация (НАЦИД)'')');

EXEC('
INSERT [dbo].[Departments] ([DepartmentId], [Alias], [Name])
VALUES (3, N''RTA'', N''Изпълнителна агенция „Автомобилна администрация“'')');

EXEC('
INSERT [dbo].[Departments] ([DepartmentId], [Alias], [Name])
VALUES (4, N''NSSI'', N''Национален осигурителен институт (НОИ)'')');

SET IDENTITY_INSERT [dbo].[Departments] OFF 


EXEC('
ALTER TABLE EserviceClients ADD DepartmentId INT NULL; ');

EXEC('
ALTER TABLE EserviceClients ADD CONSTRAINT [FK_EserviceClients_Departments]     FOREIGN KEY([DepartmentId])     REFERENCES [dbo].[Departments] ([DepartmentId]) ');


EXEC('
update EserviceClients 
set DepartmentId=1
where Alias=''Ecscs'' ');

EXEC('
update EserviceClients 
set DepartmentId=3
where Alias=''Rta'' ');

EXEC('
update EserviceClients 
set DepartmentId=2
where Alias=''Nacid'' ');

EXEC('
update EserviceClients 
set DepartmentId=4
where Alias=''ArchimedLovech'' or Alias=''ArchimedStaraZagora'' or Alias=''ArchimedHaskovo'' ');

EXEC('
--You have to add TestDepartment entry to Departments table
update EserviceClients 
set DepartmentId=99
where Alias=''TestAisClientAlias'' or Alias=''TestAisClientAlias_Borica'' ');

------------

ALTER TABLE Users ADD AccessCodeNotifications BIT NULL;

EXEC('UPDATE Users SET AccessCodeNotifications=0');

ALTER TABLE Users ALTER COLUMN AccessCodeNotifications BIT NOT NULL;

-------------

CREATE TABLE [dbo].[PaidStatusPaymentMethods](
    [PaidStatusPaymentMethodId]    INT             NOT NULL IDENTITY,
    [Alias]                     NVARCHAR(50)    NOT NULL,
    [Name]                      NVARCHAR(100)   NOT NULL,
    [Description]               NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_PaidStatusPaymentMethods] PRIMARY KEY ([PaidStatusPaymentMethodId])
);

SET IDENTITY_INSERT [PaidStatusPaymentMethods] ON

EXEC('
INSERT INTO [PaidStatusPaymentMethods]
    ([PaidStatusPaymentMethodId], [Alias], [Name], [Description])
VALUES
    (1                       , N''Other'' , N''Друг'', NULL),
    (2                       , N''CashDesk'' , N''На каса'', NULL)
');

SET IDENTITY_INSERT [PaidStatusPaymentMethods] OFF

EXEC('
ALTER TABLE PaymentRequests ADD PaidStatusPaymentMethodId INT NULL; ');

EXEC('
ALTER TABLE PaymentRequests ADD PaidStatusPaymentMethodDescription NVARCHAR(MAX) NULL; ');

EXEC('
ALTER TABLE PaymentRequests ADD CONSTRAINT [FK_PaymentRequests_PaidStatusPaymentMethods]    FOREIGN KEY ([PaidStatusPaymentMethodId])   REFERENCES [dbo].[PaidStatusPaymentMethods] ([PaidStatusPaymentMethodId]) ');


--------------------------------------------------------------
--Update database version
Update [dbo].[GlobalValues] 
SET [Value] = @DbVersion,
    [ModifyDate] = GETDATE()
WHERE [Key]='DatabaseVersion'
--Add DatabaseUpdate record
INSERT INTO [dbo].[DatabaseUpdates] (ScriptName, Notes, UpdateDate)
VALUES (@ScriptName, @Notes, GETDATE())
-------------------------------------------------------------
--End commands block
COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH







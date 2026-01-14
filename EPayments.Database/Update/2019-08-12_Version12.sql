ALTER TABLE EserviceClients ADD [IsEpayVposEnabled]	BIT NULL;
update EserviceClients set [IsEpayVposEnabled]=0;
ALTER TABLE EserviceClients ALTER COLUMN [IsEpayVposEnabled] BIT NOT NULL;

ALTER TABLE VposResults DROP CONSTRAINT [FK_VposResults_VposClients];
ALTER TABLE VposResults DROP COLUMN [VposClientId];

DELETE from VposClients where alias='Epay'
update VposClients set name=N'БОРИКА Виртуален ПОС', Alias=N'Borica' where Alias='UniCredit';

ALTER TABLE Departments DROP COLUMN Alias;


ALTER TABLE EserviceClients ALTER COLUMN [ServiceName] NVARCHAR(200) NULL;
update EserviceClients set ServiceName=NULL where ServiceName='';

ALTER TABLE EserviceClients ADD [BoricaVposRequestSignCertExpHideAdminMsg]	BIT NULL;


ALTER TABLE VposResults DROP COLUMN [VposClientId];


ALTER TABLE EserviceClients DROP COLUMN BoricaVposResponseSignCertFileName;
ALTER TABLE EserviceClients DROP COLUMN BoricaVposResponseSignCertValidTo;
ALTER TABLE EserviceClients DROP COLUMN BoricaVposResponseSignCertExpMailSend;



CREATE TABLE [dbo].[InternalAdminUsers](
    [InternalAdminUserId]                   INT                 NOT NULL IDENTITY,    
    [Name]                                  NVARCHAR(100)       NOT NULL,
	[Egn]									NVARCHAR(10)        NOT NULL UNIQUE,
	[IsSuperadmin]                          BIT                 NOT NULL,
    [IsActive]                              BIT                 NOT NULL,
	[CreateDate]                            DATETIME2           NOT NULL,
CONSTRAINT [PK_InternalAdminUsers]                              PRIMARY KEY ([InternalAdminUserId])
);


SET IDENTITY_INSERT [dbo].[InternalAdminUsers] ON 
GO
INSERT [dbo].[InternalAdminUsers] ([InternalAdminUserId], [Name], Egn, IsActive, CreateDate, IsSuperadmin)
VALUES (1, N'Демо', '0000000000', 1, GETDATE(), 1)
SET IDENTITY_INSERT [dbo].[InternalAdminUsers] OFF
GO


ALTER TABLE EserviceAdminUsers ALTER COLUMN [DepartmentId] INT NULL;


--PRODUCTION--



da se testvat:
- dsk ecomm - turizma
- borica ruse
- fibank bansko


- da se aktivira epay za vsichki klienti... tova li sa IBAN-ite???
- da se vidqt izteklite sertifikati za BORICA... vij file-a


---
- da test-vam EDelivery login - i da set-na <add key="EPayments.Web:EDeliveryAuthSecret"

- da setup-na ap, koito da ping-va jobs app-a
- !!! da se test-va i da se opravi izprashtaneto na PAYMENT notifikaciq kam AIS-ite... da se proveri dali raboti?


----------

update GlobalValues set Value='12' where [key]='DatabaseVersion'

INSERT [dbo].[DatabaseUpdates] ([ScriptName], [Notes], [UpdateDate]) 
VALUES ( N'2019-08-12_Version12', NULL, CAST(N'2019-08-12 22:15:26.6700000' AS DateTime2))
GO


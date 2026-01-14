BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '3';
DECLARE @ScriptName NVARCHAR(MAX) = '2015-11-25_Version3';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

--Update ECSCS [EserviceClients] values
UPDATE [EserviceClients] 
SET [AccountBIC] = 'STSABGSF',
    [AccountIBAN] = 'BG29STSA93000004959760'
WHERE [Alias] = 'Ecscs'

--Update VposResults table
ALTER TABLE VposResults ADD IsPaymentCanceledByUser BIT NULL;
EXEC('UPDATE VposResults SET IsPaymentCanceledByUser=0');
EXEC('ALTER TABLE VposResults ALTER COLUMN IsPaymentCanceledByUser BIT NOT NULL');
ALTER TABLE VposResults ALTER COLUMN PostUrl NVARCHAR(MAX) NOT NULL;

--Update EserviceClients table
exec sp_RENAME 'EserviceClients.[VposMerchantId]' , 'DskVposMerchantId', 'COLUMN';
exec sp_RENAME 'EserviceClients.[VposMerchantPassword]' , 'DskVposMerchantPassword', 'COLUMN';

ALTER TABLE EserviceClients ADD BoricaVposTerminalId                    NVARCHAR(100)       NULL;
ALTER TABLE EserviceClients ADD BoricaVposMerchantId                    NVARCHAR(100)       NULL;
ALTER TABLE EserviceClients ADD BoricaVposRequestSignCertFileName       NVARCHAR(MAX)       NULL;
ALTER TABLE EserviceClients ADD BoricaVposRequestSignCertPassword       NVARCHAR(MAX)       NULL;
ALTER TABLE EserviceClients ADD BoricaVposRequestSignCertValidTo        DATETIME2           NULL;
ALTER TABLE EserviceClients ADD BoricaVposRequestSignCertExpMailSend    BIT                 NULL;
ALTER TABLE EserviceClients ADD BoricaVposResponseSignCertFileName      NVARCHAR(MAX)       NULL;
ALTER TABLE EserviceClients ADD BoricaVposResponseSignCertValidTo       DATETIME2           NULL;
ALTER TABLE EserviceClients ADD BoricaVposResponseSignCertExpMailSend   BIT                 NULL;

--Insert [VposClients]
PRINT 'Insert [VposClients]'

SET IDENTITY_INSERT [dbo].[VposClients] ON 

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (2, N'УниКредит Виртуален ПОС', N'UniCredit', N'https://gatet.borica.bg/boreps/registerTransaction', 1)

SET IDENTITY_INSERT [dbo].[VposClients] OFF


--Insert [EserviceClients]
PRINT 'Insert [EserviceClients]'

SET IDENTITY_INSERT [dbo].[EserviceClients] ON 

EXEC('INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint], BoricaVposTerminalId, BoricaVposMerchantId, BoricaVposRequestSignCertFileName, BoricaVposRequestSignCertPassword, BoricaVposRequestSignCertValidTo, BoricaVposResponseSignCertFileName, BoricaVposResponseSignCertValidTo, BoricaVposResponseSignCertExpMailSend)
VALUES (2, 2, N''Rta'', N''rta_e0d07219-2cfc-4d5e-9d17-b777b7992422'', N''Портал на изпълнителна агенция "Автомобилна администрация"'', N''Промяна на списъците към разрешенията за обучение за придобиване на правоспособност за управление на МПС'', NULL, ''УниКредит Булбанк'', ''UNCRBGSF'', '''', N''SDFKLGHLRTYJH467345LKJDFVLEKJ578LEKJLK24L5JK7LM45E7KLM4356H23KH7JKL3647JKJKLH45KJLH6376782A2136NMCNHAWPASDIJETLVSKDJPWRTYEPYPASW'', 1, NULL, NULL, ''bbac2475-bb65-4d27-ba7f-440180d3363c'', 0, 0, N'''', N'''', N''RTA_ePaymentsRequestSign_Prod.p12'', '''', ''2017-11-25'', N''3DS Payment Gateway 2014.cer'', ''2016-07-28'', 0)');

EXEC('INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint], BoricaVposTerminalId, BoricaVposMerchantId, BoricaVposRequestSignCertFileName, BoricaVposRequestSignCertPassword, BoricaVposRequestSignCertValidTo, BoricaVposResponseSignCertFileName, BoricaVposResponseSignCertValidTo, BoricaVposResponseSignCertExpMailSend)
VALUES (3, 2, N''Nacid'', N''nacid_f1417e88-519a-4741-9adf-1496b798403e'', N''Портал за електронни административни услуги на "НАЦИД"'', N''Издаване на служебна бележка '', N''Издаване на служебна бележка за завършени научноизследователски проекти.'' + char(13) + char(10) + N''Издаване на служебна бележка за депозиран научен ръкопис.'' + char(13) + char(10) + N''Издаване на служебна бележка за защитен дисертационен труд.'' + char(13) + char(10) + N''Издаване на служебна бележка за заета академична длъжност.'', ''УниКредит Булбанк'', ''UNCRBGSF'', ''BG97UNCR96603119995010'', N''AEIRJBVM4673MKSDLKMPETNSIPM74H789SCDWMVBN782F34564DAQWEERERMBVD95IAPJTRRRYUUIIGVBN34566UDFBUHERIJHSXCKMV234345634LKSDFGJRTYIJ4J5'', 1, NULL, NULL, ''ef2cea14-fb93-4c4c-a059-d3ae61857285'', 0, 0, N'''', N'''', N''NACID_ePaymentsRequestSign_Prod.p12'', '''', ''2017-12-02'', N''3DS Payment Gateway 2014.cer'', ''2016-07-28'', 0)');

SET IDENTITY_INSERT [dbo].[EserviceClients] OFF

--Create VposBoricaRequests
PRINT 'VposBoricaRequests';

CREATE TABLE [dbo].[VposBoricaRequests](
    [VposBoricaRequestId]               INT                 NOT NULL IDENTITY,
    [RequestIdentifier]                 NVARCHAR(15)        NOT NULL UNIQUE,
    [PaymentRequestId]                  INT                 NOT NULL,
    [VposRedirectId]                    INT                 NULL,
    [RedirectUrl]                       NVARCHAR(MAX)       NOT NULL,
    [CreateDate]                        DATETIME            NOT NULL,
CONSTRAINT [PK_VposBoricaRequests]                          PRIMARY KEY([VposBoricaRequestId]),
CONSTRAINT [FK_VposBoricaRequests_PaymentRequests]          FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_VposBoricaRequests_VposRedirects]            FOREIGN KEY([VposRedirectId])   REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
);


--Create VposBoricaRequests
PRINT 'EventRegisterNotifications'

CREATE TABLE [dbo].[EventRegisterNotifications](
    [EventRegisterNotificationId]   INT            NOT NULL IDENTITY,
    [PaymentRequestId]              INT             NULL,
    [EventTime]                     DATETIME2       NOT NULL,
    [EventType]                     NVARCHAR(MAX)   NULL,
    [EventDescription]              NVARCHAR(MAX)   NULL,
    [EventDocRegNumber]             NVARCHAR(MAX)   NULL,
    [NotificationStatusId]          INT             NOT NULL,
    [FailedAttempts]                INT             NOT NULL,
    [FailedAttemptsErrors]          NVARCHAR(MAX)   NULL,
    [SendNotBefore]                 DATETIME2       NULL,
    [CreateDate]                    DATETIME2       NOT NULL,
    [ModifyDate]                    DATETIME2       NOT NULL,
    [Version]                       ROWVERSION      NOT NULL,
    CONSTRAINT [PK_EventRegisterNotifications]                           PRIMARY KEY CLUSTERED ([EventRegisterNotificationId] ASC),
    CONSTRAINT [FK_EventRegisterNotifications_PaymentRequests]           FOREIGN KEY ([PaymentRequestId])    REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
    CONSTRAINT [FK_EventRegisterNotifications_NotificationStatuses]      FOREIGN KEY (NotificationStatusId)  REFERENCES [dbo].[NotificationStatuses] (NotificationStatusId)
);

PRINT 'Insert [GlobalValues]'

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'EventRegisterNotificationJobLastInvocationTime', NULL, GETDATE())


--UPDATE Ecscs service client BIC and IBAN
UPDATE EserviceClients
SET AccountBIC=N'STSABGSF',
    AccountIBAN=N'BG29STSA93000004959760'
WHERE Alias=N'Ecscs'

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







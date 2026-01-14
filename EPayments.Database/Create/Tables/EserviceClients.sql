PRINT 'EserviceClients'
GO

CREATE TABLE [dbo].[EserviceClients](
    [EserviceClientId]							INT                 NOT NULL IDENTITY,
    [Gid]										UNIQUEIDENTIFIER    NOT NULL UNIQUE,
    [Alias]										NVARCHAR(100)       NOT NULL,
    [AisName]									NVARCHAR(100)       NOT NULL,
    [ServiceName]								NVARCHAR(200)       NULL,
    [ServiceDescription]						NVARCHAR(MAX)       NULL,
    [DepartmentId]								INT                 NULL,
    [AccountBank]								NVARCHAR(100)       NOT NULL,
    [AccountBIC]								NVARCHAR(50)        NOT NULL,
    [AccountIBAN]								NVARCHAR(50)        NOT NULL,
    [ClientId]									NVARCHAR(100)       NOT NULL,
    [SecretKey]									NVARCHAR(200)       NOT NULL,
    [VposClientId]								INT                 NULL,
    [DskVposMerchantId]							NVARCHAR(100)       NULL,
    [DskVposMerchantPassword]					NVARCHAR(MAX)       NULL,
    [BoricaVposTerminalId]						NVARCHAR(100)       NULL,
    [BoricaVposMerchantId]						NVARCHAR(100)       NULL,
    [BoricaVposRequestSignCertFileName]			NVARCHAR(MAX)       NULL,
    [BoricaVposRequestSignCertPassword]			NVARCHAR(MAX)       NULL,
    [BoricaVposRequestSignCertValidTo]			DATETIME2           NULL,
    [BoricaVposRequestSignCertExpMailSend]		BIT                 NULL,
	[BoricaVposRequestSignCertExpHideAdminMsg]  BIT                 NULL,
    --[BoricaVposResponseSignCertFileName]		NVARCHAR(MAX)       NULL,
    --[BoricaVposResponseSignCertValidTo]		DATETIME2           NULL,
    --[BoricaVposResponseSignCertExpMailSend]	BIT                 NULL,
	[FiBankVposAccessKeystoreFilename]			NVARCHAR(MAX)       NULL,
	[FiBankVposAccessKeystorePassword]			NVARCHAR(MAX)       NULL,
	[IsEpayVposEnabled]							BIT                 NOT NULL,
    [IsAuthPassAuthorized]						BIT                 NOT NULL,
    [IsActive]									BIT                 NOT NULL,
 CONSTRAINT [PK_EserviceClients]                        PRIMARY KEY ([EserviceClientId]),
 CONSTRAINT [FK_EserviceClients_VposClients]            FOREIGN KEY ([VposClientId])             REFERENCES [dbo].[VposClients] ([VposClientId]),
 CONSTRAINT [FK_EserviceClients_Departments]            FOREIGN KEY ([DepartmentId])             REFERENCES [dbo].[Departments] ([DepartmentId])
);

GO

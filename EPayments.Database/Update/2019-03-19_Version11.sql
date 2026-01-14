PRINT 'Insert [VposClients]'
GO

SET IDENTITY_INSERT [dbo].[VposClients] ON 
GO

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (3, N'FiBank Виртуален ПОС', N'FiBank', NULL, 1)


SET IDENTITY_INSERT [dbo].[VposClients] OFF
GO

ALTER TABLE EserviceClients ADD [FiBankVposAccessKeystoreFilename]		NVARCHAR(MAX)       NULL;
ALTER TABLE EserviceClients ADD [FiBankVposAccessKeystorePassword]		NVARCHAR(MAX)       NULL;


PRINT 'VposFiBankRequests'
GO

CREATE TABLE [dbo].[VposFiBankRequests](
    [VposFiBankRequestId]					INT                 NOT NULL IDENTITY,
    [PaymentRequestId]						INT                 NOT NULL,
	[VposRedirectId]						INT                 NULL,
	[ClientIpAddress]						NVARCHAR(50)        NOT NULL,
	[TransactionId]							NVARCHAR(MAX)       NOT NULL,
    [IsVposPostCallbackReceived]			BIT                 NOT NULL,
	[VposPostCallbackReceiveDate]			DATETIME            NULL,
	[TransactionResult]						NVARCHAR(MAX)       NULL,
	[TransactionResultReceiveDate]			DATETIME            NULL,
	[IsPaymentSuccessful]					BIT					NULL,
	[ResultStatus]							INT					NOT NULL,
	[JobCheckResultFailedAttempts]			INT					NOT NULL,
	[JobCheckResultFailedAttemptsErrors]	NVARCHAR(MAX)		NULL,
	[JobLastCheckResultDate]				DATETIME            NULL,
	[JobLastCheckResultTransactionInfo]		NVARCHAR(MAX)		NULL,
    [CreateDate]							DATETIME            NOT NULL,
CONSTRAINT [PK_VposFiBankRequests]                          PRIMARY KEY([VposFiBankRequestId]),
CONSTRAINT [FK_VposFiBankRequests_PaymentRequests]          FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_VposFiBankRequests_VposRedirects]            FOREIGN KEY([VposRedirectId])   REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
);


CREATE NONCLUSTERED INDEX IX_VposFiBankRequests_ResultStatus_CreateDate_JobCheckResultFailedAttempts_JobLastCheckResultDate
ON [dbo].[VposFiBankRequests](ResultStatus, CreateDate, JobCheckResultFailedAttempts, JobLastCheckResultDate);  

GO  



----------

update GlobalValues set Value='11' where [key]='DatabaseVersion'

INSERT [dbo].[DatabaseUpdates] ([ScriptName], [Notes], [UpdateDate]) 
VALUES ( N'2019-03-19_Version11', NULL, CAST(N'2019-03-19 22:15:26.6700000' AS DateTime2))
GO


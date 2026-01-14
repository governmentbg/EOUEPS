PRINT 'VposFiBankRequests'
GO

CREATE TABLE [dbo].[VposFiBankRequests](
    [VposFiBankRequestId]					INT                 NOT NULL IDENTITY,
    [PaymentRequestId]						INT                 NOT NULL,
	[VposRedirectId]						INT                 NULL,
	[ClientIpAddress]						NVARCHAR(50)        NOT NULL,
	[TransactionId]							NVARCHAR(100)       NOT NULL,
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


CREATE NONCLUSTERED INDEX IX_VposFiBankRequests_TransactionId
ON [dbo].[VposFiBankRequests](TransactionId);  

GO  

CREATE NONCLUSTERED INDEX IX_VposFiBankRequests_ResultStatus_CreateDate_JobCheckResultFailedAttempts_JobLastCheckResultDate
ON [dbo].[VposFiBankRequests](ResultStatus, CreateDate, JobCheckResultFailedAttempts, JobLastCheckResultDate);  

GO  

PRINT 'VposDskEcommRequests'
GO

CREATE TABLE [dbo].[VposDskEcommRequests](
    [VposDskEcommRequestId]					INT                 NOT NULL IDENTITY,
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
CONSTRAINT [PK_VposDskEcommRequests]                          PRIMARY KEY([VposDskEcommRequestId]),
CONSTRAINT [FK_VposDskEcommRequests_PaymentRequests]          FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_VposDskEcommRequests_VposRedirects]            FOREIGN KEY([VposRedirectId])   REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
);


CREATE NONCLUSTERED INDEX IX_VposDskEcommRequests_TransactionId
ON [dbo].[VposDskEcommRequests](TransactionId);  

GO  

CREATE NONCLUSTERED INDEX IX_VposDskEcommRequests_ResultStatus_CreateDate_JobCheckResultFailedAttempts_JobLastCheckResultDate
ON [dbo].[VposDskEcommRequests](ResultStatus, CreateDate, JobCheckResultFailedAttempts, JobLastCheckResultDate);  

GO  

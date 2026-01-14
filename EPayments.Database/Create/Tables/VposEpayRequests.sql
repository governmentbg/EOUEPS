PRINT 'VposEpayRequests'
GO

CREATE TABLE [dbo].[VposEpayRequests](
    [VposEpayRequestId]						INT                 NOT NULL IDENTITY,
    [PaymentRequestId]						INT                 NOT NULL,
	[VposRedirectId]						INT                 NULL,
	[ClientIpAddress]						NVARCHAR(50)        NOT NULL,
	[TransactionInvoiceNo]					NVARCHAR(100)       NOT NULL UNIQUE,
    [IsVposPostCallbackReceived]			BIT                 NOT NULL,
	[VposPostCallbackReceiveDate]			DATETIME            NULL,
	[VposPostCallbackSuccessConfirmation]	BIT					NULL,
	[VposPostCallbackAisRedirectedStatus]	NVARCHAR(50)		NULL,
	[TransactionResult]						NVARCHAR(MAX)       NULL,
	[TransactionResultReceiveDate]			DATETIME            NULL,
	[IsPaymentSuccessful]					BIT					NULL,
    [CreateDate]							DATETIME            NOT NULL,
CONSTRAINT [PK_VposEpayRequests]                          PRIMARY KEY([VposEpayRequestId]),
CONSTRAINT [FK_VposEpayRequests_PaymentRequests]          FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_VposEpayRequests_VposRedirects]            FOREIGN KEY([VposRedirectId])   REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
);


CREATE NONCLUSTERED INDEX IX_VposEpayRequests_TransactionInvoiceNo
ON [dbo].[VposEpayRequests](TransactionInvoiceNo);  

GO  

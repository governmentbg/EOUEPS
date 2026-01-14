PRINT 'BoricaTransactionPaymentRequest'
GO

CREATE TABLE [dbo].[BoricaTransactionPaymentRequest] (
	[BoricaTransactionId]													INT			NOT NULL,
	[PaymentRequestId]														INT			NOT NULL,
	CONSTRAINT [PK_BoricaTransactionPaymentRequest]										PRIMARY KEY ([BoricaTransactionId], [PaymentRequestId]),
	CONSTRAINT [FK_BoricaTransactionPaymentRequest_BoricaTransactions]					FOREIGN KEY ([BoricaTransactionId]) REFERENCES [dbo].[BoricaTransactions]([BoricaTransactionId]),
	CONSTRAINT [FK_BoricaTransactionPaymentRequest_PaymentRequests]						FOREIGN KEY ([PaymentRequestId]) REFERENCES [dbo].[PaymentRequests]([PaymentRequestId])
);
PRINT 'PaymentRequestObligationLogs'
GO

CREATE TABLE [dbo].[PaymentRequestObligationLogs] (
	[PaymentRequestObligationLogsId]								INT				IDENTITY(1, 1) NOT NULL,
	[PaymentRequestId]												INT				NOT NULL,
	[ObligationStatusId]											INT				NOT NULL,
	[ChangeDate]													DATETIME2		NOT NULL,
CONSTRAINT [PK_PaymentRequestObligationLogsId]										PRIMARY KEY CLUSTERED ([PaymentRequestObligationLogsId]),
CONSTRAINT [FK_PaymentRequestObligationLogs_PaymentRequests]						FOREIGN KEY([PaymentRequestId]) REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_PaymentRequestObligationLogs_ObligationStatuses]						FOREIGN KEY([ObligationStatusId]) REFERENCES [dbo].[ObligationStatuses] ([ObligationStatusId])
);
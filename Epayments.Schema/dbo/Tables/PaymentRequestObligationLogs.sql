CREATE TABLE [dbo].[PaymentRequestObligationLogs](
	[PaymentRequestObligationLogsId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[ObligationStatusId] [int] NOT NULL,
	[ChangeDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_PaymentRequestObligationLogsId] PRIMARY KEY CLUSTERED 
(
	[PaymentRequestObligationLogsId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[PaymentRequestObligationLogs]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequestObligationLogs_ObligationStatuses] FOREIGN KEY([ObligationStatusId])
REFERENCES [dbo].[ObligationStatuses] ([ObligationStatusId])
GO

ALTER TABLE [dbo].[PaymentRequestObligationLogs] CHECK CONSTRAINT [FK_PaymentRequestObligationLogs_ObligationStatuses]
GO
ALTER TABLE [dbo].[PaymentRequestObligationLogs]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequestObligationLogs_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[PaymentRequestObligationLogs] CHECK CONSTRAINT [FK_PaymentRequestObligationLogs_PaymentRequests]
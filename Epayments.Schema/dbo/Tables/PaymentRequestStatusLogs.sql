CREATE TABLE [dbo].[PaymentRequestStatusLogs](
	[PaymentRequestStatusLogId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[PaymentRequestStatusId] [int] NOT NULL,
	[ChangeDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_PaymentRequestStatusLogs] PRIMARY KEY CLUSTERED 
(
	[PaymentRequestStatusLogId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[PaymentRequestStatusLogs]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[PaymentRequestStatusLogs] CHECK CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequests]
GO
ALTER TABLE [dbo].[PaymentRequestStatusLogs]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequestStatuses] FOREIGN KEY([PaymentRequestStatusId])
REFERENCES [dbo].[PaymentRequestStatuses] ([PaymentRequestStatusId])
GO

ALTER TABLE [dbo].[PaymentRequestStatusLogs] CHECK CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequestStatuses]
CREATE TABLE [dbo].[EbankingAccessLogs](
	[EbankingAccessLogId] [int] IDENTITY(1,1) NOT NULL,
	[EbankingClientId] [int] NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[IpAddress] [nvarchar](50) NOT NULL,
	[AccessDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_EbankingAccessLogs] PRIMARY KEY CLUSTERED 
(
	[EbankingAccessLogId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[EbankingAccessLogs]  WITH CHECK ADD  CONSTRAINT [FK_EbankingAccessLogs_EbankingClients] FOREIGN KEY([EbankingClientId])
REFERENCES [dbo].[EbankingClients] ([EbankingClientId])
GO

ALTER TABLE [dbo].[EbankingAccessLogs] CHECK CONSTRAINT [FK_EbankingAccessLogs_EbankingClients]
GO
ALTER TABLE [dbo].[EbankingAccessLogs]  WITH CHECK ADD  CONSTRAINT [FK_EbankingAccessLogs_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[EbankingAccessLogs] CHECK CONSTRAINT [FK_EbankingAccessLogs_PaymentRequests]
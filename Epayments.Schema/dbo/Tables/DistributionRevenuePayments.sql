CREATE TABLE [dbo].[DistributionRevenuePayments](
	[DistributionRevenuePaymentId] [int] NOT NULL,
	[EserviceClientId] [int] NOT NULL,
	[DistributionRevenueId] [int] NOT NULL,
	[BoricaTransactionId] [int] NOT NULL,
	[DitribtutionError] [nvarchar](500) NULL,
 CONSTRAINT [PK_DistributionRevenuePayments] PRIMARY KEY CLUSTERED 
(
	[DistributionRevenuePaymentId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DistributionRevenuePayments]  WITH CHECK ADD  CONSTRAINT [FK_DistributionRevenuePayments_BoricaTransactions] FOREIGN KEY([BoricaTransactionId])
REFERENCES [dbo].[BoricaTransactions] ([BoricaTransactionId])
GO

ALTER TABLE [dbo].[DistributionRevenuePayments] CHECK CONSTRAINT [FK_DistributionRevenuePayments_BoricaTransactions]
GO
ALTER TABLE [dbo].[DistributionRevenuePayments]  WITH CHECK ADD  CONSTRAINT [FK_DistributionRevenuePayments_DistributionRevenues] FOREIGN KEY([DistributionRevenueId])
REFERENCES [dbo].[DistributionRevenues] ([DistributionRevenueId])
GO

ALTER TABLE [dbo].[DistributionRevenuePayments] CHECK CONSTRAINT [FK_DistributionRevenuePayments_DistributionRevenues]
GO
ALTER TABLE [dbo].[DistributionRevenuePayments]  WITH CHECK ADD  CONSTRAINT [FK_DistributionRevenuePayments_EserviceClients] FOREIGN KEY([EserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[DistributionRevenuePayments] CHECK CONSTRAINT [FK_DistributionRevenuePayments_EserviceClients]
GO
ALTER TABLE [dbo].[DistributionRevenuePayments]  WITH CHECK ADD  CONSTRAINT [FK_DistributionRevenuePayments_PaymentRequests] FOREIGN KEY([DistributionRevenuePaymentId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[DistributionRevenuePayments] CHECK CONSTRAINT [FK_DistributionRevenuePayments_PaymentRequests]
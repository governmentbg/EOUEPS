CREATE TABLE [dbo].[BoricaTransactionPaymentRequest](
	[BoricaTransactionId] [int] NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
 CONSTRAINT [PK_BoricaTransactionPaymentRequest] PRIMARY KEY CLUSTERED 
(
	[BoricaTransactionId] ASC,
	[PaymentRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[BoricaTransactionPaymentRequest]  WITH CHECK ADD  CONSTRAINT [FK_BoricaTransactionPaymentRequest_BoricaTransactions] FOREIGN KEY([BoricaTransactionId])
REFERENCES [dbo].[BoricaTransactions] ([BoricaTransactionId])
GO

ALTER TABLE [dbo].[BoricaTransactionPaymentRequest] CHECK CONSTRAINT [FK_BoricaTransactionPaymentRequest_BoricaTransactions]
GO
ALTER TABLE [dbo].[BoricaTransactionPaymentRequest]  WITH CHECK ADD  CONSTRAINT [FK_BoricaTransactionPaymentRequest_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[BoricaTransactionPaymentRequest] CHECK CONSTRAINT [FK_BoricaTransactionPaymentRequest_PaymentRequests]
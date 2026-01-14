CREATE TABLE [dbo].[VposEpayRequests](
	[VposEpayRequestId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[VposRedirectId] [int] NULL,
	[ClientIpAddress] [nvarchar](50) NOT NULL,
	[TransactionInvoiceNo] [nvarchar](100) NOT NULL,
	[IsVposPostCallbackReceived] [bit] NOT NULL,
	[VposPostCallbackReceiveDate] [datetime] NULL,
	[VposPostCallbackSuccessConfirmation] [bit] NULL,
	[VposPostCallbackAisRedirectedStatus] [nvarchar](50) NULL,
	[TransactionResult] [nvarchar](max) NULL,
	[TransactionResultReceiveDate] [datetime] NULL,
	[IsPaymentSuccessful] [bit] NULL,
	[CreateDate] [datetime] NOT NULL,
 CONSTRAINT [PK_VposEpayRequests] PRIMARY KEY CLUSTERED 
(
	[VposEpayRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[TransactionInvoiceNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[VposEpayRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposEpayRequests_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[VposEpayRequests] CHECK CONSTRAINT [FK_VposEpayRequests_PaymentRequests]
GO
ALTER TABLE [dbo].[VposEpayRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposEpayRequests_VposRedirects] FOREIGN KEY([VposRedirectId])
REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
GO

ALTER TABLE [dbo].[VposEpayRequests] CHECK CONSTRAINT [FK_VposEpayRequests_VposRedirects]
CREATE TABLE [dbo].[VposFiBankRequests](
	[VposFiBankRequestId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[VposRedirectId] [int] NULL,
	[TransactionId] [nvarchar](100) NULL,
	[IsVposPostCallbackReceived] [bit] NOT NULL,
	[VposPostCallbackReceiveDate] [datetime] NULL,
	[TransactionResult] [nvarchar](max) NULL,
	[TransactionResultReceiveDate] [datetime] NULL,
	[IsPaymentSuccessful] [bit] NULL,
	[CreateDate] [datetime] NOT NULL,
	[ResultStatus] [int] NOT NULL,
	[JobCheckResultFailedAttempts] [int] NOT NULL,
	[JobCheckResultFailedAttemptsErrors] [nvarchar](max) NULL,
	[JobLastCheckResultDate] [datetime] NULL,
	[JobLastCheckResultTransactionInfo] [nvarchar](max) NULL,
	[ClientIpAddress] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_VposFiBankRequests] PRIMARY KEY CLUSTERED 
(
	[VposFiBankRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[VposFiBankRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposFiBankRequests_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[VposFiBankRequests] CHECK CONSTRAINT [FK_VposFiBankRequests_PaymentRequests]
GO
ALTER TABLE [dbo].[VposFiBankRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposFiBankRequests_VposRedirects] FOREIGN KEY([VposRedirectId])
REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
GO

ALTER TABLE [dbo].[VposFiBankRequests] CHECK CONSTRAINT [FK_VposFiBankRequests_VposRedirects]
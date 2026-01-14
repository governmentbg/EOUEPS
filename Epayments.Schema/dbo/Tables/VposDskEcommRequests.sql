CREATE TABLE [dbo].[VposDskEcommRequests](
	[VposDskEcommRequestId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[VposRedirectId] [int] NULL,
	[ClientIpAddress] [nvarchar](50) NOT NULL,
	[TransactionId] [nvarchar](100) NOT NULL,
	[IsVposPostCallbackReceived] [bit] NOT NULL,
	[VposPostCallbackReceiveDate] [datetime] NULL,
	[TransactionResult] [nvarchar](max) NULL,
	[TransactionResultReceiveDate] [datetime] NULL,
	[IsPaymentSuccessful] [bit] NULL,
	[ResultStatus] [int] NOT NULL,
	[JobCheckResultFailedAttempts] [int] NOT NULL,
	[JobCheckResultFailedAttemptsErrors] [nvarchar](max) NULL,
	[JobLastCheckResultDate] [datetime] NULL,
	[JobLastCheckResultTransactionInfo] [nvarchar](max) NULL,
	[CreateDate] [datetime] NOT NULL,
 CONSTRAINT [PK_VposDskEcommRequests] PRIMARY KEY CLUSTERED 
(
	[VposDskEcommRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[VposDskEcommRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposDskEcommRequests_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[VposDskEcommRequests] CHECK CONSTRAINT [FK_VposDskEcommRequests_PaymentRequests]
GO
ALTER TABLE [dbo].[VposDskEcommRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposDskEcommRequests_VposRedirects] FOREIGN KEY([VposRedirectId])
REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
GO

ALTER TABLE [dbo].[VposDskEcommRequests] CHECK CONSTRAINT [FK_VposDskEcommRequests_VposRedirects]
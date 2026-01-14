CREATE TABLE [dbo].[VposResults](
	[VposResultId] [int] IDENTITY(1,1) NOT NULL,
	[Gid] [uniqueidentifier] NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[PostData] [nvarchar](max) NOT NULL,
	[PostUrl] [nvarchar](max) NOT NULL,
	[IsPaymentSuccessfull] [bit] NOT NULL,
	[RequestDate] [datetime] NOT NULL,
	[IsPaymentCanceledByUser] [bit] NOT NULL,
 CONSTRAINT [PK_VposResults] PRIMARY KEY CLUSTERED 
(
	[VposResultId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Gid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[VposResults]  WITH CHECK ADD  CONSTRAINT [FK_VposResults_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[VposResults] CHECK CONSTRAINT [FK_VposResults_PaymentRequests]
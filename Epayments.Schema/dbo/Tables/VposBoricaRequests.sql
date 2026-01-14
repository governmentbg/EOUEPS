CREATE TABLE [dbo].[VposBoricaRequests](
	[VposBoricaRequestId] [int] IDENTITY(1,1) NOT NULL,
	[RequestIdentifier] [nvarchar](15) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[VposRedirectId] [int] NULL,
	[RedirectUrl] [nvarchar](max) NOT NULL,
	[CreateDate] [datetime] NOT NULL,
 CONSTRAINT [PK_VposBoricaRequests] PRIMARY KEY CLUSTERED 
(
	[VposBoricaRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[RequestIdentifier] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[VposBoricaRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposBoricaRequests_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[VposBoricaRequests] CHECK CONSTRAINT [FK_VposBoricaRequests_PaymentRequests]
GO
ALTER TABLE [dbo].[VposBoricaRequests]  WITH CHECK ADD  CONSTRAINT [FK_VposBoricaRequests_VposRedirects] FOREIGN KEY([VposRedirectId])
REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
GO

ALTER TABLE [dbo].[VposBoricaRequests] CHECK CONSTRAINT [FK_VposBoricaRequests_VposRedirects]
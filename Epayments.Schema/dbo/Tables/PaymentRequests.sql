CREATE TABLE [dbo].[PaymentRequests](
	[PaymentRequestId] [int] IDENTITY(1,1) NOT NULL,
	[Gid] [uniqueidentifier] NOT NULL,
	[PaymentRequestXmlId] [int] NOT NULL,
	[EserviceClientId] [int] NOT NULL,
	[IsPaymentMultiple] [bit] NOT NULL,
	[ServiceProviderName] [nvarchar](max) NOT NULL,
	[ServiceProviderBank] [nvarchar](100) NOT NULL,
	[ServiceProviderBIC] [nvarchar](50) NOT NULL,
	[ServiceProviderIBAN] [nvarchar](50) NOT NULL,
	[Currency] [nvarchar](10) NOT NULL,
	[PaymentTypeCode] [nvarchar](10) NULL,
	[PaymentAmount] [decimal](18, 4) NOT NULL,
	[PaymentReason] [nvarchar](200) NOT NULL,
	[ApplicantUinTypeId] [int] NOT NULL,
	[ApplicantUin] [nvarchar](50) NOT NULL,
	[ApplicantName] [nvarchar](max) NOT NULL,
	[PaymentReferenceType] [nvarchar](50) NULL,
	[PaymentReferenceNumber] [nvarchar](50) NOT NULL,
	[PaymentReferenceDate] [datetime2](7) NOT NULL,
	[ExpirationDate] [datetime2](7) NOT NULL,
	[AdditionalInformation] [nvarchar](max) NULL,
	[AdministrativeServiceUri] [nvarchar](100) NULL,
	[AdministrativeServiceSupplierUri] [nvarchar](100) NULL,
	[AdministrativeServiceNotificationURL] [nvarchar](max) NULL,
	[PaymentRequestIdentifier] [nvarchar](50) NOT NULL,
	[PaymentRequestAccessCode] [nvarchar](10) NULL,
	[IsTempRequest] [bit] NOT NULL,
	[IsVposAuthorized] [bit] NOT NULL,
	[VposResultId] [int] NULL,
	[VposAuthorizationId] [nvarchar](max) NULL,
	[PaymentRequestStatusId] [int] NOT NULL,
	[PaymentRequestStatusChangeTime] [datetime2](7) NOT NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[PaidStatusPaymentMethodId] [int] NULL,
	[PaidStatusPaymentMethodDescription] [nvarchar](max) NULL,
	[AisPaymentId] [nvarchar](max) NULL,
	[ObligationTypeId] [int] NULL,
	[ObligationStatusId] [int] NULL,
	[InitiatorId] [int] NOT NULL,
	[RedirectUrl] [nvarchar](1000) NULL,
 [PayOrder] INT NULL, 
    CONSTRAINT [PK_PaymentRequests] PRIMARY KEY CLUSTERED 
(
	[PaymentRequestId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ__PaymentR__C51E1337F24B38C5] UNIQUE NONCLUSTERED 
(
	[Gid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ__PaymentR__DFC22872823AFF80] UNIQUE NONCLUSTERED 
(
	[PaymentRequestIdentifier] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY], 
    CONSTRAINT [FK_PaymentRequests_ObligationType] FOREIGN KEY (ObligationTypeId) REFERENCES ObligationTypes(ObligationTypeId)
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_EserviceClients] FOREIGN KEY([EserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_EserviceClients]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_EServiceClients_InitiatorId] FOREIGN KEY([InitiatorId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_EServiceClients_InitiatorId]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_ObligationStatuses] FOREIGN KEY([ObligationStatusId])
REFERENCES [dbo].[ObligationStatuses] ([ObligationStatusId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_ObligationStatuses]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_PaidStatusPaymentMethods] FOREIGN KEY([PaidStatusPaymentMethodId])
REFERENCES [dbo].[PaidStatusPaymentMethods] ([PaidStatusPaymentMethodId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_PaidStatusPaymentMethods]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_PaymentRequestStatuses] FOREIGN KEY([PaymentRequestStatusId])
REFERENCES [dbo].[PaymentRequestStatuses] ([PaymentRequestStatusId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_PaymentRequestStatuses]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_PaymentRequestXmls] FOREIGN KEY([PaymentRequestXmlId])
REFERENCES [dbo].[PaymentRequestXmls] ([PaymentRequestXmlId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_PaymentRequestXmls]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_UinTypes] FOREIGN KEY([ApplicantUinTypeId])
REFERENCES [dbo].[UinTypes] ([UinTypeId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_UinTypes]
GO
ALTER TABLE [dbo].[PaymentRequests]  WITH CHECK ADD  CONSTRAINT [FK_PaymentRequests_VposResults] FOREIGN KEY([VposResultId])
REFERENCES [dbo].[VposResults] ([VposResultId])
GO

ALTER TABLE [dbo].[PaymentRequests] CHECK CONSTRAINT [FK_PaymentRequests_VposResults]
GO
ALTER TABLE [dbo].[PaymentRequests] ADD  CONSTRAINT [DF__PaymentRe__Initi__084B3915]  DEFAULT ((1)) FOR [InitiatorId]
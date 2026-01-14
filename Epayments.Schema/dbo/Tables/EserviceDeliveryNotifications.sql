CREATE TABLE [dbo].[EserviceDeliveryNotifications](
	[EserviceNotificationId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestObligationLogsId] [int] NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[EserviceClientId] [int] NOT NULL,
	[Uniqueidentifier] [nchar](20) NULL,
	[NotificationStatusId] [int] NOT NULL,
	[FailedAttempts] [int] NOT NULL,
	[FailedAttemptsErrors] [nvarchar](max) NULL,
	[SendNotBefore] [datetime2](7) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[ModifyDate] [datetime2](7) NOT NULL,
	[Version] [timestamp] NOT NULL,
	[PersonUniqueIdentifier] [nvarchar](20) NULL,
	[ResponseCodes] [nvarchar](100) NULL,
 CONSTRAINT [PK_EserviceDeliveryNotifications] PRIMARY KEY CLUSTERED 
(
	[EserviceNotificationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_EserviceClients] FOREIGN KEY([EserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_EserviceClients]
GO
ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_NotificationStatuses] FOREIGN KEY([NotificationStatusId])
REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])
GO

ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_NotificationStatuses]
GO
ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequestObligationLogs] FOREIGN KEY([PaymentRequestObligationLogsId])
REFERENCES [dbo].[PaymentRequestObligationLogs] ([PaymentRequestObligationLogsId])
GO

ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequestObligationLogs]
GO
ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequests]
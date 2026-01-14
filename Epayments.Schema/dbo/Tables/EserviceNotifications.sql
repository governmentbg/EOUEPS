CREATE TABLE [dbo].[EserviceNotifications](
	[EserviceNotificationId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[EserviceClientId] [int] NOT NULL,
	[Url] [nvarchar](max) NOT NULL,
	[PostData] [nvarchar](max) NOT NULL,
	[NotificationStatusId] [int] NOT NULL,
	[FailedAttempts] [int] NOT NULL,
	[FailedAttemptsErrors] [nvarchar](max) NULL,
	[SendNotBefore] [datetime2](7) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[ModifyDate] [datetime2](7) NOT NULL,
	[Version] [timestamp] NOT NULL,
 CONSTRAINT [PK_EserviceNotifications] PRIMARY KEY CLUSTERED 
(
	[EserviceNotificationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceNotifications_EserviceClients] FOREIGN KEY([EserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO

ALTER TABLE [dbo].[EserviceNotifications] CHECK CONSTRAINT [FK_EserviceNotifications_EserviceClients]
GO
ALTER TABLE [dbo].[EserviceNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceNotifications_NotificationStatuses] FOREIGN KEY([NotificationStatusId])
REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])
GO

ALTER TABLE [dbo].[EserviceNotifications] CHECK CONSTRAINT [FK_EserviceNotifications_NotificationStatuses]
GO
ALTER TABLE [dbo].[EserviceNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceNotifications_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[EserviceNotifications] CHECK CONSTRAINT [FK_EserviceNotifications_PaymentRequests]
CREATE TABLE [dbo].[EventRegisterNotifications](
	[EventRegisterNotificationId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestId] [int] NULL,
	[EventTime] [datetime2](7) NOT NULL,
	[EventType] [nvarchar](max) NULL,
	[EventDescription] [nvarchar](max) NULL,
	[EventDocRegNumber] [nvarchar](max) NULL,
	[NotificationStatusId] [int] NOT NULL,
	[FailedAttempts] [int] NOT NULL,
	[FailedAttemptsErrors] [nvarchar](max) NULL,
	[SendNotBefore] [datetime2](7) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[ModifyDate] [datetime2](7) NOT NULL,
	[Version] [timestamp] NOT NULL,
 CONSTRAINT [PK_EventRegisterNotifications] PRIMARY KEY CLUSTERED 
(
	[EventRegisterNotificationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EventRegisterNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EventRegisterNotifications_NotificationStatuses] FOREIGN KEY([NotificationStatusId])
REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])
GO

ALTER TABLE [dbo].[EventRegisterNotifications] CHECK CONSTRAINT [FK_EventRegisterNotifications_NotificationStatuses]
GO
ALTER TABLE [dbo].[EventRegisterNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EventRegisterNotifications_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[EventRegisterNotifications] CHECK CONSTRAINT [FK_EventRegisterNotifications_PaymentRequests]
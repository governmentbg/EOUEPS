BEGIN TRANSACTION

BEGIN TRY

CREATE TABLE [dbo].[EserviceDeliveryNotifications](
	[EserviceNotificationId] [int] IDENTITY(1,1) NOT NULL,
	[PaymentRequestObligationLogsId] [int] NOT NULL,
	[PaymentRequestId] [int] NOT NULL,
	[EserviceClientId] [int] NOT NULL,
	[Uniqueidentifier] [nchar](20) NULL,
	[MessageText] [nvarchar](max) NOT NULL,
	[NotificationStatusId] [int] NOT NULL,
	[FailedAttempts] [int] NOT NULL,
	[FailedAttemptsErrors] [nvarchar](max) NULL,
	[SendNotBefore] [datetime2](7) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[ModifyDate] [datetime2](7) NOT NULL,
	[Version] [timestamp] NOT NULL,
 CONSTRAINT [PK_EserviceDeliveryNotifications] PRIMARY KEY CLUSTERED 
(
	[EserviceNotificationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]


ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_EserviceClients] FOREIGN KEY([EserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])


ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_EserviceClients]


ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_NotificationStatuses] FOREIGN KEY([NotificationStatusId])
REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])


ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_NotificationStatuses]


ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequestObligationLogs] FOREIGN KEY([PaymentRequestObligationLogsId])
REFERENCES [dbo].[PaymentRequestObligationLogs] ([PaymentRequestObligationLogsId])


ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequestObligationLogs]


ALTER TABLE [dbo].[EserviceDeliveryNotifications]  WITH CHECK ADD  CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])


ALTER TABLE [dbo].[EserviceDeliveryNotifications] CHECK CONSTRAINT [FK_EserviceDeliveryNotifications_PaymentRequests]





ALTER TABLE [dbo].[EserviceClients]
	ADD [DeliveryAdminstrationId]						NVARCHAR(50)		NULL,
	    [DeliveryAdministrationGuid]				    nvarchar(MAX)		NULL

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-11_Version19 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO
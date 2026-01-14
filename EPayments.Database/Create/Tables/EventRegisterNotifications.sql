PRINT 'EventRegisterNotifications'
GO

CREATE TABLE [dbo].[EventRegisterNotifications](
    [EventRegisterNotificationId]   INT            NOT NULL IDENTITY,
    [PaymentRequestId]              INT             NULL,
    [EventTime]                     DATETIME2       NOT NULL,
    [EventType]                     NVARCHAR(MAX)   NULL,
    [EventDescription]              NVARCHAR(MAX)   NULL,
    [EventDocRegNumber]             NVARCHAR(MAX)   NULL,
    [NotificationStatusId]          INT             NOT NULL,
    [FailedAttempts]                INT             NOT NULL,
    [FailedAttemptsErrors]          NVARCHAR(MAX)   NULL,
    [SendNotBefore]                 DATETIME2       NULL,
    [CreateDate]                    DATETIME2       NOT NULL,
    [ModifyDate]                    DATETIME2       NOT NULL,
    [Version]                       ROWVERSION      NOT NULL,
    CONSTRAINT [PK_EventRegisterNotifications]                           PRIMARY KEY CLUSTERED ([EventRegisterNotificationId] ASC),
    CONSTRAINT [FK_EventRegisterNotifications_PaymentRequests]           FOREIGN KEY ([PaymentRequestId])    REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
    CONSTRAINT [FK_EventRegisterNotifications_NotificationStatuses]      FOREIGN KEY (NotificationStatusId)  REFERENCES [dbo].[NotificationStatuses] (NotificationStatusId)
);

GO

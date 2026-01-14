PRINT 'EserviceNotifications'
GO

CREATE TABLE [dbo].[EserviceNotifications](
    [EserviceNotificationId]    INT             NOT NULL IDENTITY,
    [PaymentRequestId]          INT             NOT NULL,
    [EserviceClientId]          INT             NOT NULL,
    [Url]                       NVARCHAR(MAX)   NOT NULL,
    [PostData]                  NVARCHAR(MAX)   NOT NULL,
    [NotificationStatusId]      INT             NOT NULL,
    [FailedAttempts]            INT             NOT NULL,
    [FailedAttemptsErrors]      NVARCHAR(MAX)   NULL,
    [SendNotBefore]             DATETIME2       NULL,
    [CreateDate]                DATETIME2       NOT NULL,
    [ModifyDate]                DATETIME2       NOT NULL,
    [Version]                   ROWVERSION      NOT NULL,
    CONSTRAINT [PK_EserviceNotifications]                           PRIMARY KEY CLUSTERED ([EserviceNotificationId] ASC),
    CONSTRAINT [FK_EserviceNotifications_PaymentRequests]           FOREIGN KEY ([PaymentRequestId])    REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
    CONSTRAINT [FK_EserviceNotifications_EserviceClients]           FOREIGN KEY ([EserviceClientId])    REFERENCES [dbo].[EserviceClients] ([EserviceClientId]),
    CONSTRAINT [FK_EserviceNotifications_NotificationStatuses]      FOREIGN KEY (NotificationStatusId)  REFERENCES [dbo].[NotificationStatuses] (NotificationStatusId)
);

GO

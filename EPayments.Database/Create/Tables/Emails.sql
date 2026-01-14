PRINT 'Emails'
GO

CREATE TABLE [dbo].[Emails](
    [EmailId]               INT             NOT NULL IDENTITY,
    [Recipient]             NVARCHAR(100)   NOT NULL,
    [MailTemplateName]      NVARCHAR(100)   NOT NULL,
    [Context]               NVARCHAR(MAX)   NULL,
    [NotificationStatusId]  INT             NOT NULL,
    [FailedAttempts]        INT             NOT NULL,
    [FailedAttemptsErrors]  NVARCHAR(MAX)   NULL,
    [CreateDate]            DATETIME2       NOT NULL,
    [ModifyDate]            DATETIME2       NOT NULL,
    [Version]               ROWVERSION      NOT NULL,

    CONSTRAINT [PK_Emails]                  PRIMARY KEY CLUSTERED ([EmailId] ASC),
    CONSTRAINT [FK_Emails_NotificationStatuses]    FOREIGN KEY ([NotificationStatusId]) REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])
);

GO

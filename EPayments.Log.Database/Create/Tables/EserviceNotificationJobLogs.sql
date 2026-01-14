PRINT 'EserviceNotificationJobLogs'
GO

CREATE TABLE [dbo].[EserviceNotificationJobLogs](
    [EserviceNotificationJobLogId]      INT              NOT NULL IDENTITY,
    [Level]                             NVARCHAR (50)    NULL,
    [Message]                           NVARCHAR (MAX)   NULL,
    [LogDate]                           DATETIME         NULL,
    CONSTRAINT [PK_EserviceNotificationJobLogs] PRIMARY KEY ([EserviceNotificationJobLogId])
);
GO

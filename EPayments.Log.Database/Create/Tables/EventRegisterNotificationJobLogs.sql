PRINT 'EventRegisterNotificationJobLogs'
GO

CREATE TABLE [dbo].[EventRegisterNotificationJobLogs](
    [EventRegisterNotificationJobLogId]     INT              NOT NULL IDENTITY,
    [Level]                                 NVARCHAR (50)    NULL,
    [Message]                               NVARCHAR (MAX)   NULL,
    [LogDate]                               DATETIME         NULL,
    CONSTRAINT [PK_EventRegisterNotificationJobLogs] PRIMARY KEY ([EventRegisterNotificationJobLogId])
);
GO

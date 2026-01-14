PRINT 'ExpiredRequestJobLogs'
GO

CREATE TABLE [dbo].[ExpiredRequestJobLogs](
    [ExpiredRequestJobLogId]      INT              NOT NULL IDENTITY,
    [Level]                 NVARCHAR (50)    NULL,
    [Message]               NVARCHAR (MAX)   NULL,
    [LogDate]               DATETIME         NULL,
    CONSTRAINT [PK_ExpiredRequestJobLogs] PRIMARY KEY ([ExpiredRequestJobLogId])
);
GO

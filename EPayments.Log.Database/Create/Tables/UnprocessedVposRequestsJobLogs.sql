PRINT 'UnprocessedVposRequestsJobLogs'
GO

CREATE TABLE [dbo].[UnprocessedVposRequestsJobLogs](
    [UnprocessedVposRequestsJobLogId]      INT                 NOT NULL IDENTITY,
    [Level]                 NVARCHAR (50)    NULL,
    [Message]               NVARCHAR (MAX)   NULL,
    [LogDate]               DATETIME         NULL,
    CONSTRAINT [PK_UnprocessedVposRequestsJobLogs] PRIMARY KEY ([UnprocessedVposRequestsJobLogId])
);
GO

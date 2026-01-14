PRINT 'EmailJobLogs'
GO

CREATE TABLE [dbo].[EmailJobLogs](
    [EmailJobLogId]      INT                 NOT NULL IDENTITY,
    [Level]                 NVARCHAR (50)    NULL,
    [Message]               NVARCHAR (MAX)   NULL,
    [LogDate]               DATETIME         NULL,
    CONSTRAINT [PK_EmailJobLogs] PRIMARY KEY ([EmailJobLogId])
);
GO

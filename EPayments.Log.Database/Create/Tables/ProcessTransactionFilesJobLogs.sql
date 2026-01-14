PRINT 'ProcessTransactionFilesJobLogs'
GO

CREATE TABLE [dbo].[ProcessTransactionFilesJobLogs](
    [ProcessTransactionFilesJobLogId]      INT              NOT NULL IDENTITY,
    [Level]                 NVARCHAR (50)    NULL,
    [Message]               NVARCHAR (MAX)   NULL,
    [LogDate]               DATETIME         NULL,
    CONSTRAINT [PK_ProcessTransactionFilesJobLogs] PRIMARY KEY ([ProcessTransactionFilesJobLogId])
);
GO

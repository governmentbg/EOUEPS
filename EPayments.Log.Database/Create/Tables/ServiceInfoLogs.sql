PRINT 'ServiceInfoLogs'
GO

CREATE TABLE [dbo].[ServiceInfoLogs](
    [ServiceInfoLogId]      INT              NOT NULL IDENTITY,
    [Level]                 NVARCHAR (50)    NOT NULL,
    [Application]           NVARCHAR (50)    NOT NULL,
    [LogDate]               DATETIME         NULL,
    [Message]               NVARCHAR (MAX)   NULL,
    [IP]                    NVARCHAR (50)    NULL,
    [Method]                NVARCHAR (10)    NULL,
    [RawUrl]                NVARCHAR (MAX)   NULL,
    [UserAgent]             NVARCHAR (MAX)   NULL,
    [SessionId]             UNIQUEIDENTIFIER NULL,
    [RequestId]             UNIQUEIDENTIFIER NULL,
    [ElapsedMilliseconds]   BIGINT           NULL,
    [Status]                NVARCHAR (MAX)   NULL,
    CONSTRAINT [PK_ServiceInfoLogs] PRIMARY KEY ([ServiceInfoLogId])
);
GO

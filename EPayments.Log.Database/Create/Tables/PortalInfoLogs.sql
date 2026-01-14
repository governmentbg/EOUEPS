PRINT 'PortalInfoLogs'
GO

CREATE TABLE [dbo].[PortalInfoLogs](
    [PortalInfoLogId]   INT              NOT NULL IDENTITY,
    [Level]             NVARCHAR (50)    NOT NULL,
    [LogDate]           DATETIME         NULL,
    [IP]                NVARCHAR (50)    NULL,
    [RawUrl]            NVARCHAR (500)   NULL,
    [Form]              NVARCHAR (MAX)   NULL,
    [UserAgent]         NVARCHAR (200)   NULL,
    [SessionId]         NVARCHAR (50)    NULL,
    [RequestId]         UNIQUEIDENTIFIER NULL,
    [Message]           NVARCHAR (MAX)   NULL,
    CONSTRAINT [PK_PortalInfoLogs] PRIMARY KEY ([PortalInfoLogId])
);
GO

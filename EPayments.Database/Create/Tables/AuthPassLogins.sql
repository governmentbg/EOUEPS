PRINT 'AuthPassLogins'
GO

CREATE TABLE [dbo].[AuthPassLogins](
    [AuthPassLoginId]   INT                 NOT NULL IDENTITY,
    [Gid]               UNIQUEIDENTIFIER    NOT NULL UNIQUE,
    [IP]                NVARCHAR (50)       NULL,
    [PostData]          NVARCHAR (MAX)      NOT NULL,
    [LogDate]           DATETIME            NOT NULL,
 CONSTRAINT [PK_AuthPassLogins] PRIMARY KEY ([AuthPassLoginId])
);

GO
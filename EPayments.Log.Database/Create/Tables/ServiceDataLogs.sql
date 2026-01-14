PRINT 'ServiceDataLogs'
GO

CREATE TABLE [dbo].[ServiceDataLogs](
    [ServiceDataLogId]  INT              NOT NULL IDENTITY,
    [Method]            NVARCHAR (100)   NOT NULL,
    [ClientId]          NVARCHAR (200)   NULL,
    [PostData]          NVARCHAR (MAX)   NULL,
    [ResponseData]      NVARCHAR (MAX)   NULL,
    [RawUrl]            NVARCHAR (MAX)   NOT NULL,
    [RequestId]         UNIQUEIDENTIFIER NULL,
    [RemoteIpAddress]   NVARCHAR (50)    NOT NULL,
    [LogDate]           DATETIME         NOT NULL,
    CONSTRAINT [PK_ServiceDataLogs]                  PRIMARY KEY ([ServiceDataLogId]),
);
GO

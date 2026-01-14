PRINT 'DatabaseUpdates'
GO

CREATE TABLE [dbo].[DatabaseUpdates](
    [DatabaseUpdateId]     INT             IDENTITY (1, 1) NOT NULL,
    [ScriptName]           NVARCHAR (200)  NOT NULL,
    [Notes]                NVARCHAR (MAX)  NULL,
    [UpdateDate]           DATETIME2       NOT NULL,

 CONSTRAINT [PK_DatabaseUpdates] PRIMARY KEY ([DatabaseUpdateId])
);

GO



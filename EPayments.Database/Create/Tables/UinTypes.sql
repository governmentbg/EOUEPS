PRINT 'UinTypes'
GO

CREATE TABLE [dbo].[UinTypes](
    [UinTypeId]     INT             NOT NULL IDENTITY,
    [Alias]         NVARCHAR(50)    NOT NULL,
    [Name]          NVARCHAR(100)   NOT NULL,
    [Description]   NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_UinTypes] PRIMARY KEY ([UinTypeId])
);

GO
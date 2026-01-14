PRINT 'GlobalValues'
GO

CREATE TABLE [dbo].[GlobalValues](
    [GlobalValueId]     INT             IDENTITY (1, 1) NOT NULL,
    [Key]               NVARCHAR (200)  NOT NULL,
    [Value]             NVARCHAR (MAX)  NULL,
    [ModifyDate]        DATETIME2       NOT NULL,

 CONSTRAINT [PK_GlobalValues] PRIMARY KEY ([GlobalValueId])
);

GO



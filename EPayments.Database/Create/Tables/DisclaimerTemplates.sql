PRINT 'DisclaimerTemplates'
GO

CREATE TABLE [dbo].[DisclaimerTemplates](
    [DisclaimerTemplateId]  INT                 NOT NULL IDENTITY,
    [DisclaimerText]        NVARCHAR (MAX)      NOT NULL,
    [CreateDate]            DATETIME            NOT NULL,
    [IsActive]              BIT                 NOT NULL,
 CONSTRAINT [PK_DisclaimerTemplates] PRIMARY KEY ([DisclaimerTemplateId])
);

GO

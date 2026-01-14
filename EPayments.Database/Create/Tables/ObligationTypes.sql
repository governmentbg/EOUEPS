PRINT 'ObligationTypes'
GO

CREATE TABLE [dbo].[ObligationTypes] (
	[ObligationTypeId]					INT				IDENTITY,
	[Name]								VARCHAR(200)	NOT NULL,
	[IsActive]							BIT				NOT NULL DEFAULT(0), 
	CONSTRAINT [PK_ObligationTypeId]					PRIMARY KEY ([ObligationTypeId])
);
GO

CREATE INDEX IDX_ObligationTypes_Name ON [dbo].[ObligationTypes]([Name]);
GO

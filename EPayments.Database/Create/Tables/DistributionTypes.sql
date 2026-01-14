PRINT 'DistributionTypes'
GO

CREATE TABLE [dbo].[DistributionTypes] (
	[DistributionTypeId]					INT					IDENTITY,
	[Name]									NVARCHAR(50)		NOT NULL,
	CONSTRAINT [PK_DistributionTypes]							PRIMARY KEY ([DistributionTypeId])
);

GO
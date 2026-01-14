PRINT 'Insert [DistributionTypes]'
GO

SET IDENTITY_INSERT [dbo].[DistributionTypes] ON;

INSERT INTO [dbo].[DistributionTypes]
	([DistributionTypeId], [Name])
VALUES
	(1, 'С банков превод');

SET IDENTITY_INSERT [dbo].[DistributionTypes] OFF;
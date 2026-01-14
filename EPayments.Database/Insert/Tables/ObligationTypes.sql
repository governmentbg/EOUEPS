PRINT 'Insert [ObligationTypes]'
GO

SET IDENTITY_INSERT [dbo].[ObligationTypes] ON;
GO

INSERT INTO [dbo].[ObligationTypes]
	([ObligationTypeId], [Name], [IsActive])
VALUES
	(1, 'Общо задължение', 1);
GO

SET IDENTITY_INSERT [dbo].[ObligationTypes] OFF;
GO
PRINT 'Insert [DatabaseUpdates]'
GO

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-11-21_Version17', NULL, GETDATE());
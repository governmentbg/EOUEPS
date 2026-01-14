PRINT 'Insert [GlobalValues]'
GO


INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'DatabaseVersion', N'Initial', GETDATE())

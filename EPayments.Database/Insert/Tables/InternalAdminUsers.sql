PRINT 'Insert [InternalAdminUsers]'
GO

SET IDENTITY_INSERT [dbo].[InternalAdminUsers] ON 
GO

INSERT [dbo].[InternalAdminUsers] ([InternalAdminUserId], [Name], Egn, IsSuperAdmin, IsActive, CreateDate)
VALUES (1, N'Демо', '9011118326', 1, 1, GETDATE())

INSERT [dbo].[InternalAdminUsers] ([InternalAdminUserId], [Name], Egn, IsSuperAdmin, IsActive, CreateDate)
VALUES (2, N'Ilian Kostov', '8311266364', 1, 1, GETDATE())

INSERT [dbo].[InternalAdminUsers] ([InternalAdminUserId], [Name], Egn, IsSuperAdmin, IsActive, CreateDate)
VALUES (4, N'Emiliyan Evgeniev', '8405263204', 1, 1, GETDATE())

SET IDENTITY_INSERT [dbo].[InternalAdminUsers] OFF
GO

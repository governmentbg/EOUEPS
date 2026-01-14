PRINT 'Insert [EserviceAdminUsers]'
GO

SET IDENTITY_INSERT [dbo].[EserviceAdminUsers] ON 
GO

INSERT [dbo].[EserviceAdminUsers] ([EserviceAdminUserId], [Username], [PasswordHash], [PasswordSalt], [IpAddressesForAccess], [Name], [IsActive], [InsufficientAmountNotifications],[OverpaidAmountNotifications],[ReferencedNotifications], [NotReferencedNotifications], [ReferringEserviceClientId])
VALUES (1, N'emil', N'AKCM5mChIKu1ztNnQ9hZ4r9PdIwOdy3HoyTeHHsWsN+1Ot98WuDcuJMl2Qi7l8tm5g==', N'1urqFSoAUT6ldJKSqYFn4A==', NULL, N'Емил Дечев Денчовски', 1, 0, 0, 0, 0, 1)
GO

SET IDENTITY_INSERT [dbo].[EserviceAdminUsers] OFF
GO

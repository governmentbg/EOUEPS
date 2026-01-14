PRINT 'Insert [EserviceAdminUserBankAccounts]'
GO

SET IDENTITY_INSERT [dbo].[EserviceAdminUserBankAccounts] ON 
GO

INSERT [dbo].[EserviceAdminUserBankAccounts] ([EserviceAdminUserBankAccountId], [EserviceAdminUserId], [EserviceBankAccountId], [IsActive])
VALUES (1, 1, 1, 1)
GO

SET IDENTITY_INSERT [dbo].[EserviceAdminUserBankAccounts] OFF
GO

PRINT 'Insert [EserviceBankAccounts]'
GO

SET IDENTITY_INSERT [dbo].[EserviceBankAccounts] ON 
GO

INSERT [dbo].[EserviceBankAccounts] ([EserviceBankAccountId], [BankId], [Iban], [IsActive], [UploadTransactions], [TransactionsFilesPathUnread], [TransactionsFilesPathRead])
VALUES (1, 2, N'BG97UNCR96603119995010', 1, 1, N'D:\Projects Files\EPayments\UniCredit transaction files\New Test\Unread', N'D:\Projects Files\EPayments\UniCredit transaction files\New Test\Read')
GO

SET IDENTITY_INSERT [dbo].[EserviceBankAccounts] OFF
GO

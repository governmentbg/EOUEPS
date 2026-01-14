PRINT 'Insert [Banks]'
GO

SET IDENTITY_INSERT [dbo].[Banks] ON 
GO

INSERT [dbo].[Banks] ([BankId], [Name], [BIC], [CertificateId])
VALUES (1, N'Банка ДСК', N'STSABGSF', NULL)
GO

INSERT [dbo].[Banks] ([BankId], [Name], [BIC], [CertificateId])
VALUES (2, N'УниКредит Булбанк', N'UNCRBGSF', NULL)
GO

SET IDENTITY_INSERT [dbo].[Banks] OFF
GO

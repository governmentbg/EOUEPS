CREATE TABLE [dbo].[EserviceAdminUserBankAccounts](
	[EserviceAdminUserBankAccountId] [int] IDENTITY(1,1) NOT NULL,
	[EserviceAdminUserId] [int] NOT NULL,
	[EserviceBankAccountId] [int] NOT NULL,
	[IsActive] [bit] NOT NULL,
 CONSTRAINT [PK_EserviceAdminUserBankAccounts] PRIMARY KEY CLUSTERED 
(
	[EserviceAdminUserBankAccountId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceAdminUserBankAccounts]  WITH CHECK ADD  CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceAdminUsers] FOREIGN KEY([EserviceAdminUserId])
REFERENCES [dbo].[EserviceAdminUsers] ([EserviceAdminUserId])
GO

ALTER TABLE [dbo].[EserviceAdminUserBankAccounts] CHECK CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceAdminUsers]
GO
ALTER TABLE [dbo].[EserviceAdminUserBankAccounts]  WITH CHECK ADD  CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceBankAccounts] FOREIGN KEY([EserviceBankAccountId])
REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
GO

ALTER TABLE [dbo].[EserviceAdminUserBankAccounts] CHECK CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceBankAccounts]
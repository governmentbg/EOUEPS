PRINT 'EserviceAdminUserBankAccounts'
GO

CREATE TABLE [dbo].[EserviceAdminUserBankAccounts](
    [EserviceAdminUserBankAccountId]            INT                 NOT NULL IDENTITY,
    [EserviceAdminUserId]                       INT                 NOT NULL,
    [EserviceBankAccountId]                     INT                 NOT NULL,
    [IsActive]                                  BIT                 NOT NULL,
CONSTRAINT [PK_EserviceAdminUserBankAccounts]                               PRIMARY KEY ([EserviceAdminUserBankAccountId]),
CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceAdminUsers]            FOREIGN KEY ([EserviceAdminUserId])     REFERENCES [dbo].[EserviceAdminUsers] ([EserviceAdminUserId]),
CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceBankAccounts]          FOREIGN KEY ([EserviceBankAccountId])   REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
);




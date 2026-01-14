PRINT 'EserviceBankAccounts'
GO

CREATE TABLE [dbo].[EserviceBankAccounts](
    [EserviceBankAccountId]                 INT                 NOT NULL IDENTITY,
    [BankId]                                INT                 NOT NULL,
    [Iban]                                  NVARCHAR(50)        NOT NULL UNIQUE,
    [UploadTransactions]                    BIT                 NOT NULL,
    [TransactionsFilesPathUnread]           NVARCHAR(MAX)       NULL,
    [TransactionsFilesPathRead]             NVARCHAR(MAX)       NULL,
    [IsActive]                              BIT                 NOT NULL,
CONSTRAINT [PK_EserviceBankAccounts]                        PRIMARY KEY ([EserviceBankAccountId]),
CONSTRAINT [FK_EserviceBankAccounts_Banks]                  FOREIGN KEY ([BankId])     REFERENCES [dbo].[Banks] ([BankId])
);




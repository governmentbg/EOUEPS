PRINT 'TransactionFiles'
GO

CREATE TABLE [dbo].[TransactionFiles](
    [TransactionFileId]             INT             NOT NULL IDENTITY,
    [EserviceBankAccountId]         INT             NOT NULL,
    [FileName]                      NVARCHAR(MAX)   NULL,
    [BankStatementIban]             NVARCHAR(100)   NOT NULL,
    [BankStatementDate]             DATETIME2       NULL,
    [BankStatementNumber]           NVARCHAR(100)   NULL,
    [CreateDate]                    DATETIME2       NOT NULL,
CONSTRAINT [PK_TransactionFiles]                           PRIMARY KEY ([TransactionFileId]),
CONSTRAINT [FK_TransactionFiles_EserviceBankAccounts]      FOREIGN KEY ([EserviceBankAccountId])        REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
);

GO

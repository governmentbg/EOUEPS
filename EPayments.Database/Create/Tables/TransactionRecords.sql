PRINT 'TransactionRecords'
GO

CREATE TABLE [dbo].[TransactionRecords](
    [TransactionRecordId]                   INT             NOT NULL IDENTITY,
    [TransactionFileId]                     INT             NOT NULL,
    [TransactionDate]                       DATETIME2       NULL,
    [TransactionAccountingDate]             DATETIME2       NULL,
    [TransactionAmount]                     DECIMAL(18, 4)  NULL,
    [TransactionReferenceId]                NVARCHAR(100)   NULL,
    [InfoSystemTransactionType]             NVARCHAR(100)   NULL,
    [InfoSystemTransactionDesc]             NVARCHAR(MAX)   NULL,
    [InfoPaymentType]                       NVARCHAR(100)   NULL,
    [InfoDocumentType]                      NVARCHAR(100)   NULL,
    [InfoDocumentNumber]                    NVARCHAR(100)   NULL,
    [InfoDocumentDate]                      DATETIME2       NULL,
    [InfoDocumentNumberDate]                NVARCHAR(150)   NULL,
    [InfoPaymentPeriodBegining]             DATETIME2       NULL,
    [InfoPaymentPeriodEnd]                  DATETIME2       NULL,
    [InfoDebtorBulstat]                     NVARCHAR(100)   NULL,
    [InfoDebtorEgn]                         NVARCHAR(100)   NULL,
    [InfoDebtorLnch]                        NVARCHAR(100)   NULL,
    [InfoDebtorName]                        NVARCHAR(300)   NULL,
    [InfoDebtorBulstatEgnLnch]              NVARCHAR(300)   NULL,
    [InfoDebtorBulstatEgnLnchName]          NVARCHAR(600)   NULL,
    [InfoAC1AuthorizationCode]              NVARCHAR(100)   NULL,
    [InfoAC1BankCardInfo]                   NVARCHAR(100)   NULL,
    [InfoPaymentDetailsRaw]                 NVARCHAR(MAX)   NULL,
    [InfoPaymentReason]                     NVARCHAR(MAX)   NULL,
    [InfoSenderBic]                         NVARCHAR(100)   NULL,
    [InfoSenderIban]                        NVARCHAR(100)   NULL,
    [InfoSenderName]                        NVARCHAR(MAX)   NULL,
    [InfoSenderIbanName]                    NVARCHAR(MAX)   NULL,
    [TransactionRecordPaymentMethodId]      INT             NOT NULL,
    [TransactionRecordRefStatusId]          INT             NOT NULL,
    [PaymentRequestId]                      INT             NULL,

CONSTRAINT [PK_TransactionRecords]                                      PRIMARY KEY ([TransactionRecordId]),
CONSTRAINT [FK_TransactionRecords_TransactionFiles]                     FOREIGN KEY ([TransactionFileId])                   REFERENCES [dbo].[TransactionFiles] ([TransactionFileId]),
CONSTRAINT [FK_TransactionRecords_TransactionRecordPaymentMethods]      FOREIGN KEY ([TransactionRecordPaymentMethodId])    REFERENCES [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId]),
CONSTRAINT [FK_TransactionRecords_TransactionRecordRefStatuses]         FOREIGN KEY ([TransactionRecordRefStatusId])        REFERENCES [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId]),
CONSTRAINT [FK_TransactionRecords_PaymentRequests]                      FOREIGN KEY ([PaymentRequestId])                    REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
);

GO

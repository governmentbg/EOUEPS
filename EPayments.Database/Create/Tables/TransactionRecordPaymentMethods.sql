PRINT 'TransactionRecordPaymentMethods'
GO

CREATE TABLE [dbo].[TransactionRecordPaymentMethods](
    [TransactionRecordPaymentMethodId]  INT             NOT NULL IDENTITY,
    [Alias]                             NVARCHAR(50)    NOT NULL,
    [Name]                              NVARCHAR(100)   NOT NULL,
    [Description]                       NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_TransactionRecordPaymentMethods] PRIMARY KEY ([TransactionRecordPaymentMethodId])
);

GO

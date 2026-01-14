PRINT 'PaymentRequestIdentifiers'
GO

CREATE TABLE [dbo].[PaymentRequestIdentifiers](
    [PaymentRequestIdentifierId]    INT             NOT NULL IDENTITY,
    [Date]                          DATETIME2       NOT NULL,
    [Counter]                       INT             NOT NULL,
 CONSTRAINT [PK_PaymentRequestIdentifiers] PRIMARY KEY ([PaymentRequestIdentifierId]),
 CONSTRAINT [UQ_PaymentRequestIdentifiers_Date_Counter]     UNIQUE     ([Date], [Counter])
);

GO

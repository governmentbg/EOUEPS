PRINT 'PaymentRequestStatuses'
GO

CREATE TABLE [dbo].[PaymentRequestStatuses](
    [PaymentRequestStatusId]    INT             NOT NULL IDENTITY,
    [Alias]                     NVARCHAR(50)    NOT NULL,
    [Name]                      NVARCHAR(100)   NOT NULL,
    [Description]               NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_PaymentRequestStatuses] PRIMARY KEY ([PaymentRequestStatusId])
);

GO

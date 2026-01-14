PRINT 'PaidStatusPaymentMethods'
GO

CREATE TABLE [dbo].[PaidStatusPaymentMethods](
    [PaidStatusPaymentMethodId]    INT             NOT NULL IDENTITY,
    [Alias]                     NVARCHAR(50)    NOT NULL,
    [Name]                      NVARCHAR(100)   NOT NULL,
    [Description]               NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_PaidStatusPaymentMethods] PRIMARY KEY ([PaidStatusPaymentMethodId])
);

GO

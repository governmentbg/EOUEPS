PRINT 'PaymentRequestXmls'
GO

CREATE TABLE [dbo].[PaymentRequestXmls](
    [PaymentRequestXmlId]   INT             NOT NULL IDENTITY,
    [RequestContent]        NVARCHAR(MAX)   NOT NULL,
    [ResponseContent]       NVARCHAR(MAX)   NOT NULL,
    [IsRequestAccepted]     BIT             NOT NULL,
    [CreateDate]            DATETIME2       NOT NULL,
 CONSTRAINT [PK_PaymentRequestXmls] PRIMARY KEY ([PaymentRequestXmlId])
);

GO

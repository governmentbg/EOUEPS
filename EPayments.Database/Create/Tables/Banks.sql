PRINT 'Banks'
GO

CREATE TABLE [dbo].[Banks](
    [BankId]            INT             NOT NULL IDENTITY,
    [Name]              NVARCHAR(200)   NOT NULL,
    [BIC]               NVARCHAR(50)    NOT NULL,
    [CertificateId]     INT             NULL,
 CONSTRAINT [PK_Banks] PRIMARY KEY ([BankId])
);

GO

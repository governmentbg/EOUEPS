PRINT 'EbankingClients'
GO

CREATE TABLE [dbo].[EbankingClients](
    [EbankingClientId]      INT             NOT NULL IDENTITY,
    [Name]                  NVARCHAR(200)   NOT NULL,
    [Description]           NVARCHAR(MAX)   NULL,
    [ClientId]              NVARCHAR(100)   NOT NULL,
    [SecretKey]             NVARCHAR(200)   NOT NULL,
    [IsActive]              BIT             NOT NULL,
 CONSTRAINT [PK_EbankingClients] PRIMARY KEY ([EbankingClientId]),
);

GO

PRINT 'VposClients'
GO

CREATE TABLE [dbo].[VposClients](
    [VposClientId]      INT             NOT NULL IDENTITY,
    [Name]              NVARCHAR(200)   NOT NULL,
    [Alias]             NVARCHAR(100)   NOT NULL,
    [PaymentRequestUrl] NVARCHAR(MAX)   NULL,
    [IsActive]          BIT             NOT NULL,
 CONSTRAINT [PK_VposClients] PRIMARY KEY ([VposClientId]),
);

GO

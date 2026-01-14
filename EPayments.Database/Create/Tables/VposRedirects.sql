PRINT 'VposRedirects'
GO

CREATE TABLE [dbo].[VposRedirects](
    [VposRedirectId]                    INT                 NOT NULL IDENTITY,
    [Gid]                               UNIQUEIDENTIFIER    NOT NULL UNIQUE,
    [PaymentRequestIdentifier]          NVARCHAR(50)        NULL,
    [OkUrl]                             NVARCHAR(MAX)       NULL,
    [CancelUrl]                         NVARCHAR(MAX)       NULL,
    [IP]                                NVARCHAR (50)       NULL,
    [LogDate]                           DATETIME            NOT NULL,
CONSTRAINT [PK_VposRedirects]                             PRIMARY KEY([VposRedirectId]),
);





PRINT 'Certificates'
GO

CREATE TABLE [dbo].[Certificates](
    [CertificateId]             INT                 NOT NULL IDENTITY,
    [CertificateFile]           VARBINARY (MAX)     NULL,
    [CertificateSubject]        NVARCHAR (MAX)      NULL,
    [CertificateIssuer]         NVARCHAR (MAX)      NULL,
    [CertificateThumbprint]     NVARCHAR (200)      NULL,
    CONSTRAINT [PK_Certificates] PRIMARY KEY ([CertificateId]),
    CONSTRAINT [UK_Certificates_Thumbprint] UNIQUE NONCLUSTERED ([CertificateThumbprint] ASC)
);

GO

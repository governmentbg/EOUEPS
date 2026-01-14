PRINT 'LoginAttemptLogs'
GO

CREATE TABLE [dbo].[LoginAttemptLogs](
    [LoginAttemptLogId]     INT                 NOT NULL IDENTITY,
    [LogDate]               DATETIME2(7)        NOT NULL,
    [IP]                    NVARCHAR (50)       NOT NULL,
    [CertificateFile]       VARBINARY (MAX)     NULL,
    [CertificateIssuer]     NVARCHAR (MAX)      NULL,
    [CertificatePolicies]   NVARCHAR (MAX)      NULL,
    [CertificateSubject]    NVARCHAR (MAX)      NULL,
    [AlternativeSubject]    NVARCHAR (MAX)      NULL,
    [CertificateThumbprint] NVARCHAR (200)      NULL,
    [Egn]                   NVARCHAR (50)       NULL,
    [NameLatin]             NVARCHAR (250)      NULL,
    [ErrorCode]             NVARCHAR (30)       NULL,
    [IsIisErrorOccurred]    BIT                 NOT NULL,
    [IsUesParsed]           BIT                 NOT NULL,
    [IsPersonal]            BIT                 NULL,
    [IsEgnParsed]           BIT                 NULL,
    [IsNameParsed]          BIT                 NULL,
    [IsLoginSuccessful]     BIT                 NOT NULL,

 CONSTRAINT [PK_LoginAttemptLogs] PRIMARY KEY ([LoginAttemptLogId])
);

GO



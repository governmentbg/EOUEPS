PRINT 'Users'
GO

CREATE TABLE [dbo].[Users](
    [UserId]                       INT              NOT NULL IDENTITY,
    [Egn]                           NVARCHAR(10)    NOT NULL UNIQUE,
    [Email]                         NVARCHAR(100)   NULL,
    [RequestNotifications]          BIT             NOT NULL,
    [StatusNotifications]           BIT             NOT NULL,
    [AccessCodeNotifications]       BIT             NOT NULL,
    [LastCertificateId]         INT                 NULL,
    [LastDisclaimerTemplateId]  INT                 NULL,
    CONSTRAINT [PK_Users] PRIMARY KEY ([UserId]),
    CONSTRAINT [FK_Users_Certificates] FOREIGN KEY ([LastCertificateId]) REFERENCES [dbo].[Certificates] ([CertificateId]),
    CONSTRAINT [FK_Users_DisclaimerTemplates] FOREIGN KEY ([LastDisclaimerTemplateId]) REFERENCES [dbo].[DisclaimerTemplates] ([DisclaimerTemplateId]),
    CONSTRAINT [UK_Users_Egn] UNIQUE NONCLUSTERED ([Egn] ASC)
);

GO

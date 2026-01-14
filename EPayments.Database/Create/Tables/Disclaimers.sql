PRINT 'Disclaimers'
GO

CREATE TABLE [dbo].[Disclaimers](
    [DisclaimerId]          INT            NOT NULL IDENTITY,
    [UserId]                INT            NOT NULL,
    [CertificateId]         INT            NOT NULL,
    [DisclaimerTemplateId]  INT            NOT NULL,
    [SignedXml]             NVARCHAR (MAX) NOT NULL,
    [SignDate]              DATETIME       NOT NULL,
    CONSTRAINT [PK_Disclaimers]                         PRIMARY KEY ([DisclaimerId] ASC),
    CONSTRAINT [FK_Disclaimers_Certificates]            FOREIGN KEY ([CertificateId])           REFERENCES [dbo].[Certificates] ([CertificateId]),
    CONSTRAINT [FK_Disclaimers_DisclaimerTemplates]     FOREIGN KEY ([DisclaimerTemplateId])    REFERENCES [dbo].[DisclaimerTemplates] ([DisclaimerTemplateId]),
    CONSTRAINT [FK_Disclaimers_Users]                   FOREIGN KEY ([UserId])                REFERENCES [dbo].[Users] ([UserId])
);

GO

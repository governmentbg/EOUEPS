PRINT 'VposBoricaRequests'
GO

CREATE TABLE [dbo].[VposBoricaRequests](
    [VposBoricaRequestId]               INT                 NOT NULL IDENTITY,
    [RequestIdentifier]                 NVARCHAR(15)        NOT NULL UNIQUE,
    [PaymentRequestId]                  INT                 NOT NULL,
    [VposRedirectId]                    INT                 NULL,
    [RedirectUrl]                       NVARCHAR(MAX)       NOT NULL,
    [CreateDate]                        DATETIME            NOT NULL,
CONSTRAINT [PK_VposBoricaRequests]                          PRIMARY KEY([VposBoricaRequestId]),
CONSTRAINT [FK_VposBoricaRequests_PaymentRequests]          FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_VposBoricaRequests_VposRedirects]            FOREIGN KEY([VposRedirectId])   REFERENCES [dbo].[VposRedirects] ([VposRedirectId])
);





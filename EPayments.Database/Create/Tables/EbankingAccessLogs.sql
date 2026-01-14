PRINT 'EbankingAccessLogs'
GO

CREATE TABLE [dbo].[EbankingAccessLogs](
    [EbankingAccessLogId]                   INT                 NOT NULL IDENTITY,
    [EbankingClientId]                      INT                 NOT NULL,
    [PaymentRequestId]                      INT                 NOT NULL,
    [IpAddress]                             NVARCHAR (50)       NOT NULL,
    [AccessDate]                            DATETIME2           NOT NULL,
CONSTRAINT [PK_EbankingAccessLogs]                             PRIMARY KEY ([EbankingAccessLogId]),
CONSTRAINT [FK_EbankingAccessLogs_EbankingClients]             FOREIGN KEY ([EbankingClientId])            REFERENCES [dbo].[EbankingClients] ([EbankingClientId]),
CONSTRAINT [FK_EbankingAccessLogs_PaymentRequests]             FOREIGN KEY ([PaymentRequestId])            REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
);




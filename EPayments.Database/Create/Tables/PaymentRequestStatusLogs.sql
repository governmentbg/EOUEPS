PRINT 'PaymentRequestStatusLogs'
GO

CREATE TABLE [dbo].[PaymentRequestStatusLogs](
    [PaymentRequestStatusLogId]     INT         NOT NULL IDENTITY ,
    [PaymentRequestId]              INT         NOT NULL,
    [PaymentRequestStatusId]        INT         NOT NULL,
    [ChangeDate]                    DATETIME2   NOT NULL,
CONSTRAINT [PK_PaymentRequestStatusLogs]                            PRIMARY KEY ([PaymentRequestStatusLogId]),
CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequests]            FOREIGN KEY([PaymentRequestId])         REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_PaymentRequestStatusLogs_PaymentRequestStatuses]     FOREIGN KEY([PaymentRequestStatusId])   REFERENCES [dbo].[PaymentRequestStatuses] ([PaymentRequestStatusId])
);

GO
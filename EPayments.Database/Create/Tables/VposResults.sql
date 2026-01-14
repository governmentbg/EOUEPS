PRINT 'VposResults'
GO

CREATE TABLE [dbo].[VposResults](
    [VposResultId]                       INT                 NOT NULL IDENTITY,
    [Gid]                                UNIQUEIDENTIFIER    NOT NULL UNIQUE,
    [PaymentRequestId]                   INT                 NOT NULL,
    [PostData]                           NVARCHAR(MAX)       NOT NULL,
    [PostUrl]                            NVARCHAR(MAX)       NOT NULL,
    [IsPaymentSuccessfull]               BIT                 NOT NULL,
    [IsPaymentCanceledByUser]            BIT                 NOT NULL,
    [RequestDate]                        DATETIME            NOT NULL,
CONSTRAINT [PK_VposResults]                             PRIMARY KEY([VposResultId]),
CONSTRAINT [FK_VposResults_PaymentRequests]             FOREIGN KEY([PaymentRequestId])     REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
);

ALTER TABLE [dbo].[PaymentRequests] ADD CONSTRAINT [FK_PaymentRequests_VposResults] FOREIGN KEY([VposResultId]) REFERENCES [dbo].[VposResults] ([VposResultId]);




PRINT 'Insert PaymentRequestStatuses'
GO

SET IDENTITY_INSERT [PaymentRequestStatuses] ON

INSERT INTO [PaymentRequestStatuses]
    ([PaymentRequestStatusId], [Alias], [Name], [Description])
VALUES
    (1                       , N'Pending' , N'Очаква плащане', NULL),
    (2                       , N'Authorized' , N'Платено с карта', NULL),
    (3                       , N'Ordered' , N'Платено по банков път', NULL),
    (4                       , N'Paid' , N'Получено плащане', NULL),
    (5                       , N'Expired' , N'Изтекъл срок', NULL),
    (6                       , N'Canceled' , N'Отказано', NULL),
    (7                       , N'Suspended' , N'Прекратена услуга', NULL)
GO

SET IDENTITY_INSERT [PaymentRequestStatuses] OFF

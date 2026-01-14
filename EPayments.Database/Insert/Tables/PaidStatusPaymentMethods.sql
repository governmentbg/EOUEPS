PRINT 'Insert PaidStatusPaymentMethods'
GO

SET IDENTITY_INSERT [PaidStatusPaymentMethods] ON

INSERT INTO [PaidStatusPaymentMethods]
    ([PaidStatusPaymentMethodId], [Alias], [Name], [Description])
VALUES
    (1                       , N'Other' , N'Друг', NULL),
    (2                       , N'CashDesk' , N'На каса', NULL)
GO

SET IDENTITY_INSERT [PaidStatusPaymentMethods] OFF

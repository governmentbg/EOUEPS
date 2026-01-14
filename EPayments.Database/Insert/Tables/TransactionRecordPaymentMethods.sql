PRINT 'Insert TransactionRecordPaymentMethods'
GO

SET IDENTITY_INSERT [TransactionRecordPaymentMethods] ON

INSERT INTO [TransactionRecordPaymentMethods]
    ([TransactionRecordPaymentMethodId], [Alias], [Name], [Description])
VALUES
    (1                       , N'BankOrder' , N'По банков път', NULL),
    (2                       , N'POS' , N'Физически ПОС', NULL),
    (3                       , N'VPOS' , N'Виртуален ПОС', NULL)
GO

SET IDENTITY_INSERT [TransactionRecordPaymentMethods] OFF

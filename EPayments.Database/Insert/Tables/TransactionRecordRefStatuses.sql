PRINT 'Insert TransactionRecordRefStatuses'
GO

SET IDENTITY_INSERT [TransactionRecordRefStatuses] ON

INSERT INTO [TransactionRecordRefStatuses]
    ([TransactionRecordRefStatusId], [Alias], [Name], [Description])
VALUES
    (1                       , N'NotReferenced' , N'Друго', NULL),
    (2                       , N'ReferencedSuccessfully' , N'Платено задължение', NULL),
    (3                       , N'ReferencedInsufficientAmount' , N'Недостатъчна сума', NULL),
    (4                       , N'ReferencedOverpaidAmount' , N'Надплатена сума', NULL)
GO

SET IDENTITY_INSERT [TransactionRecordRefStatuses] OFF

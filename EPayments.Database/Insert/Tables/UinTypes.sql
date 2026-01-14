PRINT 'Insert UinTypes'
GO

SET IDENTITY_INSERT [UinTypes] ON

INSERT INTO [UinTypes]
    ([UinTypeId], [Alias], [Name], [Description])
VALUES
    (1          , N'Egn' , N'EГН', NULL),
    (2          , N'Lnch' , N'ЛНЧ', NULL),
    (3          , N'Bulstat' , N'Булстат', NULL)
GO

SET IDENTITY_INSERT [UinTypes] OFF

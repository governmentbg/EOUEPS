PRINT 'Create [TestTransactionResults]'
GO

CREATE TABLE [dbo].[TestTransactionResults](
	[TestTransactionResultId] [int] IDENTITY(1,1) NOT NULL,
	[Id] [nvarchar](100) NULL,
	[PostData] [nvarchar](max) NULL,
    [QueryStringParams] [nvarchar](max) NULL,
    [Description] [nvarchar](max) NULL,
    [LogDate] DATETIME2(7) NULL,
 CONSTRAINT [PK_TestTransactionResults] PRIMARY KEY CLUSTERED 
(
	[TestTransactionResultId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


PRINT 'Insert test [EserviceClients]'
GO

SET IDENTITY_INSERT [dbo].[EserviceClients] ON 
GO

--INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint])
--VALUES (101, 1, N'TestAisClientAlias', N'testAisClient', N'Тестов АИС', N'Тестовo електронна услуга', NULL, 'Тестова УниKредит Булбанк', 'ТестBIC', 'ТестIBAN', N'8F70C29ACBB38F39B0900C26B3A20B0683E62C97A4E578358904686D0023988D0C9873AD23EE87003B36EE5221617FEC0345E3B1138FE1B57EF5DE4771E3CF42', 1, N'444466', N'', '71f1199e-fa51-433c-8dda-e0dee07c3a3e', 0, 0)

--INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DepartmentId], IsEpayVposEnabled)
--VALUES (101, 1, N'TestAisClientAlias', N'testAisClient', N'Тестов АИС', N'Тестовo електронна услуга', NULL, 'Райфайзенбанк', 'RZBBBGSF', 'BG94RZBB91551060362319', N'8F70C29ACBB38F39B0900C26B3A20B0683E62C97A4E578358904686D0023988D0C9873AD23EE87003B36EE5221617FEC0345E3B1138FE1B57EF5DE4771E3CF42', 1, N'90000003', N'test3test3', '71f1199e-fa51-433c-8dda-e0dee07c3a3e', 0, 99, 1)

--INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], BoricaVposTerminalId, BoricaVposMerchantId, BoricaVposRequestSignCertFileName, BoricaVposRequestSignCertPassword, BoricaVposRequestSignCertValidTo, [DepartmentId], IsEpayVposEnabled)
--VALUES (102, 2, N'TestAisClientAlias_Borica', N'testAisClient_Borica', N'Тестов АИС_Borica', N'Тестовo електронна услуга_Borica', NULL, 'ОББ', 'UBBSBGSF', 'BG94RZBB91551060362319', N'KLJSHLK45H6LK56J7234LKJDFLKBM5P0789IKFGBMP4ASYL357893278CVBCVCNSDFSFGYFU76834RBSDZSADASUIPEVXT85MOP23456CBERERTYRTUPN2136MBNCYYK', 1, null, null, '273df0a4-dac7-4720-8d75-017a6f51af2b', 0, N'62160735', N'6210015673', N'RTA_ePaymentsRequestSign_Test.p12', '', '2030-11-25', 99, 1)

SET IDENTITY_INSERT [dbo].[EserviceClients] OFF
GO


SET IDENTITY_INSERT [dbo].[EbankingClients] ON 
GO

INSERT [dbo].[EbankingClients] ([EbankingClientId], [ClientId], [Name], [Description], [SecretKey], [IsActive])
VALUES (1, N'testBankClient', N'Тестов клиент за банки', NULL, N'4E78C4A8564B9204B9D7A06DD799A286F6260742F72700F77631FCCBC0C9497F3E2FF1F6B3DFD5EA7E2224EDCCCE37889B0D9F3AEF5F7E4E69CA7E20675BD255', 1)
GO

SET IDENTITY_INSERT [dbo].[EbankingClients] OFF
GO



PRINT 'Insert [PaymentRequestXmls]'
GO

SET IDENTITY_INSERT [PaymentRequestXmls] ON

INSERT INTO [PaymentRequestXmls]
    ([PaymentRequestXmlId], [RequestContent], [ResponseContent], [CreateDate], [IsRequestAccepted])
VALUES
    (1          , N'<?xml version="1.0" encoding="utf-8"?>
<PaymentRequest xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://ereg.egov.bg/segment/R-0015">
    <PaymentTypeCode>1</PaymentTypeCode>
    <Currency>Currency</Currency>
    <PaymentAmount>40</PaymentAmount>
    <ElectronicAdministrativeServiceHeader>
        <ElectronicServiceProviderBasicData xmlns="http://ereg.egov.bg/segment/0009-000152">
            <EntityBasicData xmlns="http://ereg.egov.bg/segment/0009-000002">
                <Name xmlns="http://ereg.egov.bg/segment/0009-000013">Български институт по метрология</Name>
                <Identifier xmlns="http://ereg.egov.bg/segment/0009-000013">175092070</Identifier>
            </EntityBasicData>
            <ElectronicServiceProviderType xmlns="http://ereg.egov.bg/segment/0009-000002">ElectronicServiceProviderType</ElectronicServiceProviderType>
        </ElectronicServiceProviderBasicData>
    </ElectronicAdministrativeServiceHeader>
    <ElectronicServiceProviderBankAccount>
        <BIC xmlns="http://ereg.egov.bg/segment/R-0030">BIC</BIC>
        <IBAN xmlns="http://ereg.egov.bg/segment/R-0030">IBAN</IBAN>
        <EntityBasicData xmlns="http://ereg.egov.bg/segment/R-0030">
            <Name xmlns="http://ereg.egov.bg/segment/0009-000013">BankName</Name>
        </EntityBasicData>
    </ElectronicServiceProviderBankAccount>
    <ElectronicServiceApplicant>
        <RecipientGroup xmlns="http://ereg.egov.bg/segment/0009-000016">
            <Recipient>
                <Entity xmlns="http://ereg.egov.bg/segment/0009-000015">
                    <Name xmlns="http://ereg.egov.bg/segment/0009-000013">НИК-ФИ ООД</Name>
                    <Identifier xmlns="http://ereg.egov.bg/segment/0009-000013">130025035</Identifier>
                </Entity>
            </Recipient>
        </RecipientGroup>
    </ElectronicServiceApplicant>
    <PaymentReason>Такса за едминистративна услуга</PaymentReason>
    <PaymentReference>
        <PaymentReferenceType xmlns="http://ereg.egov.bg/segment/R-0029">PaymentReferenceType</PaymentReferenceType>
        <PaymentReferenceNumber xmlns="http://ereg.egov.bg/segment/R-0029">PaymentReferenceNumber</PaymentReferenceNumber>
        <PaymentReferenceDate xmlns="http://ereg.egov.bg/segment/R-0029">2015-07-16T15:43:54.8312123+03:00</PaymentReferenceDate>
    </PaymentReference>
    <PaymentPeriod>
        <PaymentPeriodFromDate xmlns="http://ereg.egov.bg/segment/R-0008">2015-07-16T15:43:54.8312123+03:00</PaymentPeriodFromDate>
        <PaymentPeriodToDate xmlns="http://ereg.egov.bg/segment/R-0008">2015-07-16T15:43:54.8312123+03:00</PaymentPeriodToDate>
    </PaymentPeriod>
    <PaymentRequestExpirationDate>2015-07-16T15:43:54.8312123+03:00</PaymentRequestExpirationDate>
    <AdditionalInformationInPaymentRequest>AdditionalInformationInPaymentRequest</AdditionalInformationInPaymentRequest>
    <ElectronicAdministrativeServiceUriSUNAU>ElectronicAdministrativeServiceUriSUNAU</ElectronicAdministrativeServiceUriSUNAU>
    <ElectronicAdministrativeServiceSupplierUriRA>ElectronicAdministrativeServiceSupplierUriRA</ElectronicAdministrativeServiceSupplierUriRA>
    <ElectronicAdministrativeServiceNotificationURL>ElectronicAdministrativeServiceNotificationURL</ElectronicAdministrativeServiceNotificationURL>
</PaymentRequest>' , N'', GETDATE(), 1),
    (2          , N'<?xml version="1.0" encoding="utf-8"?>
<PaymentRequest xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://ereg.egov.bg/segment/R-0015">
    <PaymentTypeCode>1</PaymentTypeCode>
    <Currency>Currency</Currency>
    <PaymentAmount>40</PaymentAmount>
    <ElectronicAdministrativeServiceHeader>
        <ElectronicServiceProviderBasicData xmlns="http://ereg.egov.bg/segment/0009-000152">
            <EntityBasicData xmlns="http://ereg.egov.bg/segment/0009-000002">
                <Name xmlns="http://ereg.egov.bg/segment/0009-000013">Български институт по метрология</Name>
                <Identifier xmlns="http://ereg.egov.bg/segment/0009-000013">175092070</Identifier>
            </EntityBasicData>
            <ElectronicServiceProviderType xmlns="http://ereg.egov.bg/segment/0009-000002">ElectronicServiceProviderType</ElectronicServiceProviderType>
        </ElectronicServiceProviderBasicData>
    </ElectronicAdministrativeServiceHeader>
    <ElectronicServiceProviderBankAccount>
        <BIC xmlns="http://ereg.egov.bg/segment/R-0030">BIC</BIC>
        <IBAN xmlns="http://ereg.egov.bg/segment/R-0030">IBAN</IBAN>
        <EntityBasicData xmlns="http://ereg.egov.bg/segment/R-0030">
            <Name xmlns="http://ereg.egov.bg/segment/0009-000013">BankName</Name>
        </EntityBasicData>
    </ElectronicServiceProviderBankAccount>
    <ElectronicServiceApplicant>
        <RecipientGroup xmlns="http://ereg.egov.bg/segment/0009-000016">
            <Recipient>
                <Entity xmlns="http://ereg.egov.bg/segment/0009-000015">
                    <Name xmlns="http://ereg.egov.bg/segment/0009-000013">НИК-ФИ ООД</Name>
                    <Identifier xmlns="http://ereg.egov.bg/segment/0009-000013">130025035</Identifier>
                </Entity>
            </Recipient>
        </RecipientGroup>
    </ElectronicServiceApplicant>
    <PaymentReason>Такса за едминистративна услуга</PaymentReason>
    <PaymentReference>
        <PaymentReferenceType xmlns="http://ereg.egov.bg/segment/R-0029">PaymentReferenceType</PaymentReferenceType>
        <PaymentReferenceNumber xmlns="http://ereg.egov.bg/segment/R-0029">PaymentReferenceNumber</PaymentReferenceNumber>
        <PaymentReferenceDate xmlns="http://ereg.egov.bg/segment/R-0029">2015-07-16T15:43:54.8312123+03:00</PaymentReferenceDate>
    </PaymentReference>
    <PaymentPeriod>
        <PaymentPeriodFromDate xmlns="http://ereg.egov.bg/segment/R-0008">2015-07-16T15:43:54.8312123+03:00</PaymentPeriodFromDate>
        <PaymentPeriodToDate xmlns="http://ereg.egov.bg/segment/R-0008">2015-07-16T15:43:54.8312123+03:00</PaymentPeriodToDate>
    </PaymentPeriod>
    <PaymentRequestExpirationDate>2015-07-16T15:43:54.8312123+03:00</PaymentRequestExpirationDate>
    <AdditionalInformationInPaymentRequest>AdditionalInformationInPaymentRequest</AdditionalInformationInPaymentRequest>
    <ElectronicAdministrativeServiceUriSUNAU>ElectronicAdministrativeServiceUriSUNAU</ElectronicAdministrativeServiceUriSUNAU>
    <ElectronicAdministrativeServiceSupplierUriRA>ElectronicAdministrativeServiceSupplierUriRA</ElectronicAdministrativeServiceSupplierUriRA>
    <ElectronicAdministrativeServiceNotificationURL>ElectronicAdministrativeServiceNotificationURL</ElectronicAdministrativeServiceNotificationURL>
</PaymentRequest>' , N'', GETDATE(), 1),
    (3          , N'' , N'', GETDATE(), 1),
    (4          , N'' , N'', GETDATE(), 1),
    (5          , N'' , N'', GETDATE(), 1),
    (6          , N'' , N'', GETDATE(), 1),
    (7          , N'' , N'', GETDATE(), 1),

    (8          , N'' , N'', GETDATE(), 1),
    (9          , N'' , N'', GETDATE(), 1),
    (10          , N'' , N'', GETDATE(), 1),
    (11          , N'' , N'', GETDATE(), 1),
    (12          , N'' , N'', GETDATE(), 1),
    (13          , N'' , N'', GETDATE(), 1),
    (14          , N'' , N'', GETDATE(), 1),
    (15          , N'' , N'', GETDATE(), 1),
    (16          , N'' , N'', GETDATE(), 1),
    (17          , N'' , N'', GETDATE(), 1),

    (18          , N'' , N'', GETDATE(), 1),
    (19          , N'' , N'', GETDATE(), 1),
    (20          , N'' , N'', GETDATE(), 1),
    (21          , N'' , N'', GETDATE(), 1),
    (22          , N'' , N'', GETDATE(), 1),
    (23          , N'' , N'', GETDATE(), 1),
    (24          , N'' , N'', GETDATE(), 1),
    (25          , N'' , N'', GETDATE(), 1),
    (26          , N'' , N'', GETDATE(), 1),
    (27          , N'' , N'', GETDATE(), 1)
GO


PRINT 'Insert [PaymentRequestXmls]'
GO

SET IDENTITY_INSERT [PaymentRequestXmls] OFF

SET IDENTITY_INSERT [dbo].[PaymentRequestIdentifiers] ON 
GO

PRINT 'Insert [PaymentRequestIdentifiers]'
GO

INSERT [dbo].[PaymentRequestIdentifiers] ([PaymentRequestIdentifierId], [Date], [Counter]) VALUES (1, CAST(0x070000000000373A0B AS DateTime2), 7)

SET IDENTITY_INSERT [dbo].[PaymentRequestIdentifiers] OFF
GO


PRINT 'Insert [Users]'
GO

INSERT [dbo].Users ([Egn], Email, RequestNotifications, StatusNotifications, AccessCodeNotifications)
VALUES ('130025035', NULL, 0, 0, 0)

INSERT [dbo].Users ([Egn], Email, RequestNotifications, StatusNotifications, AccessCodeNotifications)
VALUES ('9011118326', NULL, 0, 0, 0)


PRINT 'Insert [PaymentRequests]'
GO

SET IDENTITY_INSERT [dbo].[PaymentRequests] ON 
GO

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (1, 1, 1, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(45.0000 AS Decimal(18, 4)), N'Заплащане на такса за семестър', 1, N'7101268870', N'Марина Илиева Иванова', N'1', N'24', '2017-01-27 00:00:00.0000000', '2017-07-27 00:00:00.0000000', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507211', NULL, 3, GETDATE(), '1e63da78-4361-4c25-b94b-e0fda1fd65fe', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (2, 2, 2, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(12.0000 AS Decimal(18, 4)), N'Такса за удостоверение', 1, N'000194046', N'ТЪРГОВСКО ПРОМИШЛЕНА ПАЛАТА ВР', N'1', N'43A', '2017-01-27 00:00:00.0000000', '2017-07-27 00:00:00.0000000', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507212', NULL, 3, GETDATE(), '44925d30-e685-479e-b998-c864dc64efaa', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (3, 1, 3, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(175.0000 AS Decimal(18, 4)), N'Удостоверение APOSTILLE', 1, N'121284099', 'NТБС ТРАДИКС ООД', N'1', N'183', '2017-01-27 00:00:00.0000000', '2017-07-27 00:00:00.0000000', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507213', NULL, 1, GETDATE(), '139b091b-9776-4bc3-81b9-8730fa1539fc', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (4, 1, 4, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(45.0000 AS Decimal(18, 4)), N'Патентен данък', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№222415', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507214', NULL, 1, GETDATE(), '0e8aa973-10b7-481d-8d6d-5a922be37d48', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest], [PaymentRequestAccessCode])
VALUES (5, 1, 5, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(112.0000 AS Decimal(18, 4)), N'Годишен данък върху недвижимите имоти', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№724858', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507215', NULL, 1, GETDATE(), 'b9cea6ca-686b-4966-80d2-553716cdc5e4', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0, '1234567890')

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (6, 1, 6, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(148.0000 AS Decimal(18, 4)), N'Данък върху общата годишна данъчна основа', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№464658', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507216', NULL, 4, GETDATE(), '40f70c13-2542-438c-b11d-a04b4ddef3e0', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (7, 1, 7, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(262.2800 AS Decimal(18, 4)), N'Годишен данък МПС', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№624688', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507217', NULL, 4, GETDATE(), '5447f4d7-f26e-4a95-8a82-7713044de975', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)




INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (8, 1, 8, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(154.0000 AS Decimal(18, 4)), N'Данък върху общата годишна данъчна основа', 3, N'130025035', N'НИК-ФИ ООД', N'1', N'№122788', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507218', NULL, 1, GETDATE(), '1d48798b-6b66-474a-a4d9-0a89709cc0fb', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (9, 1, 9, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(654.2400 AS Decimal(18, 4)), N'Корпоративен данък', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№51923', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507219', NULL, 1, GETDATE(), 'b252c4e2-214b-48df-a671-41435e8f7202', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (10, 1, 10, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(48.0000 AS Decimal(18, 4)), N'Данък при придобиване на имущества по дарение', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№932452', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507220', NULL, 1, GETDATE(), 'd82cb8ec-929a-4f7a-9519-3d63ad4edc50', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (11, 1, 11, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(29.9900 AS Decimal(18, 4)), N'Патентен данък', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№222115', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507221', NULL, 1, GETDATE(), 'df6fc678-894a-406a-b3a3-2c08723a1203', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (12, 1, 12, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(112.0000 AS Decimal(18, 4)), N'Годишен данък върху недвижимите имоти', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№722858', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507222', NULL, 1, GETDATE(), '8b08c330-ffcc-4cdb-bdcb-35af8be1eb85', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (13, 1, 13, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(154.0000 AS Decimal(18, 4)), N'Данък върху общата годишна данъчна основа', 3, N'130025035', N'НИК-ФИ ООД', N'1', N'№144788', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507223', NULL, 1, GETDATE(), '9bce8aff-3f33-4741-b2ad-d96e4b57f0db', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (14, 1, 14, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(654.2400 AS Decimal(18, 4)), N'Корпоративен данък', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№51443', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507224', NULL, 1, GETDATE(), '6a39428e-d6d7-4c32-a948-a40cf697bbcb', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (15, 1, 15, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(48.0000 AS Decimal(18, 4)), N'Данък при придобиване на имущества по дарение', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№933552', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507225', NULL, 1, GETDATE(), '538f1c5c-88b0-41ab-91f6-acfd9bed835c', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (16, 1, 16, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(29.9900 AS Decimal(18, 4)), N'Патентен данък', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№222445', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507226', NULL, 1, GETDATE(), '942c9aef-e74e-47e4-89fb-eebe586deccf', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)

INSERT [dbo].[PaymentRequests] ([PaymentRequestId], [EserviceClientId], [PaymentRequestXmlId], [ServiceProviderName], [ServiceProviderBIC], [ServiceProviderIBAN], [Currency], [PaymentAmount], [PaymentReason], [ApplicantUinTypeId], [ApplicantUin], [ApplicantName], [PaymentReferenceType], [PaymentReferenceNumber], [PaymentReferenceDate], [ExpirationDate], [AdditionalInformation], [AdministrativeServiceUri], [AdministrativeServiceSupplierUri], [AdministrativeServiceNotificationURL], [PaymentRequestIdentifier], [VposAuthorizationId], [PaymentRequestStatusId], [CreateDate], [Gid], [ServiceProviderBank], [PaymentTypeCode], [PaymentRequestStatusChangeTime], [IsPaymentMultiple], [IsVposAuthorized], [IsTempRequest])
VALUES (17, 1, 17, N'НАЦИД', N'UNCRBGSF ', N'BG97UNCR96603119995010', N'BGN', CAST(112.0000 AS Decimal(18, 4)), N'Годишен данък върху недвижимите имоти', 1, N'9011118326', N'Емил Дечев Денчовски', N'1', N'№722859', GETDATE(), '2017-01-01 00:00:00', NULL, N'ServiceUri', N'SupplierUri', NULL, N'1507227', NULL, 1, GETDATE(), 'ec8ed963-b386-4f29-b2c9-8ca746bca12d', N'УниKредит Булбанк', '1', GETDATE(), 0, 0, 0)


GO
SET IDENTITY_INSERT [dbo].[PaymentRequests] OFF
GO




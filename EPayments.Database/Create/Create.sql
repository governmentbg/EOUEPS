USE [$(dbName)]
GO

---------------------------------------------------------------
--Tables
---------------------------------------------------------------
:r $(rootPath)"\Tables\VposClients.sql"
:r $(rootPath)"\Tables\Departments.sql"
:r $(rootPath)"\Tables\Banks.sql"
:r $(rootPath)"\Tables\EserviceBankAccounts.sql"
:r $(rootPath)"\Tables\EserviceClients.sql"
:r $(rootPath)"\Tables\EbankingClients.sql"
:r $(rootPath)"\Tables\PaymentRequestStatuses.sql"
:r $(rootPath)"\Tables\PaidStatusPaymentMethods.sql"
:r $(rootPath)"\Tables\UinTypes.sql"
:r $(rootPath)"\Tables\PaymentRequestXmls.sql"
:r $(rootPath)"\Tables\PaymentRequests.sql"
:r $(rootPath)"\Tables\PaymentRequestStatusLogs.sql"
:r $(rootPath)"\Tables\PaymentRequestIdentifiers.sql"
:r $(rootPath)"\Tables\EbankingAccessLogs.sql"
:r $(rootPath)"\Tables\Certificates.sql"
:r $(rootPath)"\Tables\DisclaimerTemplates.sql"
:r $(rootPath)"\Tables\Users.sql"
:r $(rootPath)"\Tables\Disclaimers.sql"
:r $(rootPath)"\Tables\LoginAttemptLogs.sql"
:r $(rootPath)"\Tables\NotificationStatuses.sql"
:r $(rootPath)"\Tables\Emails.sql"
:r $(rootPath)"\Tables\EserviceNotifications.sql"
:r $(rootPath)"\Tables\EventRegisterNotifications.sql"
:r $(rootPath)"\Tables\EserviceAdminUsers.sql"
:r $(rootPath)"\Tables\EserviceAdminUserBankAccounts.sql"
:r $(rootPath)"\Tables\TransactionRecordRefStatuses.sql"
:r $(rootPath)"\Tables\TransactionRecordPaymentMethods.sql"
:r $(rootPath)"\Tables\TransactionFiles.sql"
:r $(rootPath)"\Tables\TransactionRecords.sql"
:r $(rootPath)"\Tables\VposResults.sql"
:r $(rootPath)"\Tables\VposRedirects.sql"
:r $(rootPath)"\Tables\VposBoricaRequests.sql"
:r $(rootPath)"\Tables\VposFiBankRequests.sql"
:r $(rootPath)"\Tables\VposEpayRequests.sql"
:r $(rootPath)"\Tables\VposDskEcommRequests.sql"
:r $(rootPath)"\Tables\AuthPassLogins.sql"
:r $(rootPath)"\Tables\GlobalValues.sql"
:r $(rootPath)"\Tables\DatabaseUpdates.sql"
:r $(rootPath)"\Tables\InternalAdminUsers.sql"


---------------------------------------------------------------
--Procedures
---------------------------------------------------------------
:r $(rootPath)"\Procedures\spGetPaymentRequestCounter.sql"

--Triggers
---------------------------------------------------------------
:r $(rootPath)"\Triggers\trForInsertPaymentRequest.sql"
:r $(rootPath)"\Triggers\trForUpdatePaymentRequest.sql"
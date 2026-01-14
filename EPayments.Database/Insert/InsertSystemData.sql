USE [$(dbName)]
GO

---------------------------------------------------------------
--Inserts
---------------------------------------------------------------

:r $(rootPath)"\Tables\PaymentRequestStatuses.sql"
:r $(rootPath)"\Tables\PaidStatusPaymentMethods.sql"
:r $(rootPath)"\Tables\UinTypes.sql"
:r $(rootPath)"\Tables\NotificationStatuses.sql"
:r $(rootPath)"\Tables\VposClients.sql"
:r $(rootPath)"\Tables\Departments.sql"
:r $(rootPath)"\Tables\EserviceClients.sql"
:r $(rootPath)"\Tables\GlobalValues.sql"

:r $(rootPath)"\Tables\Banks.sql"

:r $(rootPath)"\Tables\EserviceAdminUsers.sql"
:r $(rootPath)"\Tables\EserviceBankAccounts.sql"
:r $(rootPath)"\Tables\EserviceAdminUserBankAccounts.sql"

:r $(rootPath)"\Tables\TransactionRecordRefStatuses.sql"
:r $(rootPath)"\Tables\TransactionRecordPaymentMethods.sql"

:r $(rootPath)"\Tables\InternalAdminUsers.sql"




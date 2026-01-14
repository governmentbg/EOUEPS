USE [$(dbName)]
GO

---------------------------------------------------------------
--Tables
---------------------------------------------------------------
:r $(rootPath)"\Tables\ServiceInfoLogs.sql"
:r $(rootPath)"\Tables\ServiceDataLogs.sql"

:r $(rootPath)"\Tables\PortalInfoLogs.sql"

:r $(rootPath)"\Tables\EmailJobLogs.sql"
:r $(rootPath)"\Tables\EserviceNotificationJobLogs.sql"
:r $(rootPath)"\Tables\EventRegisterNotificationJobLogs.sql"
:r $(rootPath)"\Tables\ExpiredRequestJobLogs.sql"
:r $(rootPath)"\Tables\ProcessTransactionFilesJobLogs.sql"
:r $(rootPath)"\Tables\UnprocessedVposRequestsJobLogs.sql"

:r $(rootPath)"\Tables\GlobalValues.sql"
:r $(rootPath)"\Tables\DatabaseUpdates.sql"

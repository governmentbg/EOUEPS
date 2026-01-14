SET QUOTED_IDENTIFIER ON
GO

PRINT '------ Creating EPayments'

:setvar rootPath ".\Create"
:r $(rootPath)"\CreateDB.sql"
:r $(rootPath)"\Create.sql"

:setvar rootPath ".\Insert"
:r $(rootPath)"\InsertSystemData.sql"
:r $(rootPath)"\InsertTestPaymentRequests.sql"
:r $(rootPath)"\InsertTransactionsData.sql"

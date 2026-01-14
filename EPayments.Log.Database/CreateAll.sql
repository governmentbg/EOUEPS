SET QUOTED_IDENTIFIER ON
GO

PRINT '------ Creating EpaymentsLogs'
:setvar rootPath ".\Create"
:r $(rootPath)"\CreateDB.sql"
:r $(rootPath)"\Create.sql"
:setvar rootPath ".\Insert"
:r $(rootPath)"\InsertSystemData.sql"

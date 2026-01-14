SET QUOTED_IDENTIFIER ON
GO

PRINT '------ Migrating EPayments'

:setvar rootPath ".\Update"
:r $(rootPath)"\2020-10-28_Version12.9.sql"
:r $(rootPath)"\2020-10-28_Version13.sql"
:r $(rootPath)"\2020-11-03_Version14.sql"
:r $(rootPath)"\2020-11-06_Version15.sql"
:r $(rootPath)"\2020-11-11_Version16.sql"
:r $(rootPath)"\2020-11-21_Version17.sql"
:r $(rootPath)"\2020-12-02_Version18.sql"
:r $(rootPath)"\2020-12-09_Version19.sql"
:r $(rootPath)"\2020-12-11_Version20.sql"
:r $(rootPath)"\2020-12-15_Version21.sql"
:r $(rootPath)"\2020-12-18_Version22.sql"
:r $(rootPath)"\2020-12-21_Version23.sql"
:r $(rootPath)"\2020-12-23_Version24.sql"
:r $(rootPath)"\2021-01-04_Version25.sql"
:r $(rootPath)"\2021-01-07_Version26.sql"
:r $(rootPath)"\2021-03-02_Version27.sql"
:r $(rootPath)"\2021-05-15_Version28.sql"
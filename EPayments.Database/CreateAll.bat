@sqlcmd -S "." -D "EPayments" -v dbName="EPayments" -i "CreateAll.sql"
@sqlcmd -S "." -D "EPayments" -v dbName="EPayments" -i "Migrate.sql"
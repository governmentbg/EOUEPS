PRINT 'spGetPaymentRequestCounter'
GO

CREATE PROCEDURE [dbo].[spGetPaymentRequestCounter] @Date DATETIME2
AS
BEGIN

declare @currentDate DATETIME2 = (select top 1 DATEADD(dd, DATEDIFF(dd, 0, getdate()), 0));
declare @dateRecordExist int = (select count(*) from PaymentRequestIdentifiers where [Date]=@currentDate)

if (@dateRecordExist = 0)
	begin
		insert into PaymentRequestIdentifiers ([Date], [Counter]) values (@currentDate, 0)
	end

update [dbo].PaymentRequestIdentifiers set [Counter] = [Counter] + 1
output inserted.[Counter]
where [Date]=@currentDate
 
END

GO

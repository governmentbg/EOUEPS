CREATE   PROCEDURE [dbo].[spGetPaymentRequestTotal]
AS
BEGIN

DECLARE @identifier INT = (SELECT TOP 1 [TransactionRequestIdentifierId] FROM [TransactionRequestIdentifiers] WHERE [counter] < 999999 ORDER BY [TransactionRequestIdentifierId] DESC);

IF @identifier IS NULL
BEGIN
	INSERT INTO [TransactionRequestIdentifiers]
	([Counter])
	VALUES
	(0);
END

UPDATE [dbo].[TransactionRequestIdentifiers] 
	SET [Counter] = [Counter] + 1
	OUTPUT inserted.[Counter]
	WHERE [TransactionRequestIdentifierId] = (SELECT MAX([TransactionRequestIdentifierId]) FROM [TransactionRequestIdentifiers])
END
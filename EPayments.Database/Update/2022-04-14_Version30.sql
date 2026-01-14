BEGIN TRANSACTION

BEGIN TRY

INSERT INTO [dbo].[PaymentRequestStatuses] 
	([PaymentRequestStatusId], [Alias], [Name], [Description]) 
VALUES 
	( 9, 'InProcess', 'В процес на обработка', NULL);

ALTER TABLE [dbo].[PaymentRequest]
	ADD [ObligationTypeId] INT NOT NULL;


ALTER TABLE [dbo].[ObligationTypes]
	ADD 
		[PaymentTypeCode] NVARCHAR(50) NULL,
		[AlgorithmId] INT NOT NULL;


UPDATE [dbo].[PaymentRequest] 
    SET Value = 8 
	WHERE 
	[Key] = 'ObligationTypeId'; 

---- Insert Values in [ObligationTypes]
--SET IDENTITY_INSERT [dbo].[ObligationTypes] ON;

--INSERT INTO [dbo].[ObligationTypes]
--	([ObligationTypeId], [Name], [PaymentTypeCode], [AlgorithmId], [IsActive])
--VALUES
--	(1, 'Такси за административни услуги', '448007', 1, 1),
--	(2, 'Данък върху недвижимите имоти', '442100', 1, 1),
--	(3, 'Данък върху наследствата', '442200', 1, 1),
--	(4, 'Данък върху превозните средства', '442300', 1, 1),
--	(5, 'Такси за битови отпадъци', '442400', 1, 1),
--	(6, 'Данък при придобиване на имущества по дарения и възмезден начин', '442500', 1, 1),
--	(7, 'Туристически данък', '442800', 1, 1),
--	(8, 'Други данъци', '443400', 1, 1),
--	(9, 'Вноски от приходи на общински предприятия и институции', '443700', 1, 1),
--	(10, 'Нетни приходи от продажби на услуги, стоки и продукция', '444000', 1, 1),
--	(11, 'Приходи от наеми на имущество', '441400', 1, 1),
--	(12, 'Приходи от наеми на земя', '444100', 1, 1),
--	(13, 'Приходи от лихви по текущи банкови сметки', '444300', 1, 1),
--	(14, 'Приходи от лихви по срочни депозити', '443400', 1, 1),
--	(15, 'Дивидент', '444800', 1, 1),
--	(16, 'Конфискувани средства и приходи от продажби на конфискувани вещи', '444900', 1, 1),
--	(17, 'Дарения,помощи и други безвъзмездно получени суми от страната', '445100', 1, 1),
--	(18, 'Дарения,помощи и други безвъзмездно получени суми от чужбина', '445200', 1, 1),
--	(19, 'Приходи от продажби на дълготрайни материални активи', '445500', 1, 1),
--	(20, 'Приходи от продажби на земя', '445600', 1, 1),
--	(21, 'Приходи от концесии', '445700', 1, 1),
--	(22, 'Приходи от ликвидиране на общински предприятия', '445800', 1, 1),
--	(23, 'Приходи от продажба на нематериални активи', '445900', 1, 1),
--	(24, 'Глоби, санкции и неустойки,нак. лихви, обезщетения и начети', '446500', 1, 1),
--	(25, 'Други неданъчни приходи', '447000', 1, 1),
--	(26, 'Такси за технически услуги', '448001', 1, 1),
--	(27, 'Такси за ползване на детски градини', '448002', 1, 1),
--	(28, 'Такси за ползване на детски ясли и други по здравеопазването', '448003', 1, 1),
--	(29, 'Такси за ползване на детски лагери и други по социалния отдих', '448004', 1, 1),
--	(30, 'Такси за ползване на домашен социален патронаж и други социални услуги', '448005', 1, 1),
--	(31, 'Окончателен годишен (патентен) данък и данък върху таксиметров превоз на пътници', '441400', 0, 1),
--	(32, 'Такси за ползване на пазари, тържища и др.', '448008', 1, 1),
--	(33, 'Такси за ползване на полудневни детски градини', '448009', 1, 1),
--	(34, 'Такси за ползване на общежития и други по образованието', '448010', 1, 1),
--	(35, 'Такси за гробни места', '448011', 1, 1),
--	(36, 'Такси за притежаване на куче', '448013', 1, 1),
--	(37, 'Други общински такси', '448090', 1, 1),
--	(38, 'Приходи от такси за общините по Закона за опазване на околната среда', '448081', 1, 1),
--	(39, 'Неразпределени общински приходи за операции по чл. 4, ал. 1 от ЗОПБ', '448844', 1, 1),
--	(40, 'Неидентифицирани преводи по сметки 7311 (т. 21 от ДДС № 06/2006 г. )', '448868', 1, 1),
--	(41, 'Преводи по с/ки 8 4 с многоредов документ', '448888', 1, 1),
--	(42, 'дебитно - контролираща сметка - код за централизация към сметка 7304', '449944', 1, 1),
--	(43, 'дебитно - контролираща сметка - код за преводи от сметка 7304', '449998', 1, 1)

--SET IDENTITY_INSERT [dbo].[ObligationTypes] OFF;

-- Update [ObligationTypes]
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за административни услуги',[PaymentTypeCode] = '448007', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 1;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Данък върху недвижимите имоти',[PaymentTypeCode] = '442100', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 2;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Данък върху наследствата',[PaymentTypeCode] = '442200', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 3;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Данък върху превозните средства',[PaymentTypeCode] = '442300', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 4;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за битови отпадъци',[PaymentTypeCode] = '442400', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 5;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Данък при придобиване на имущества по дарения и възмезден начин',[PaymentTypeCode] = '442500', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 6;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Туристически данък',[PaymentTypeCode] = '442800', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 7;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Други данъци',[PaymentTypeCode] = '443400', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 8;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Вноски от приходи на общински предприятия и институции',[PaymentTypeCode] = '443700', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 9;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Нетни приходи от продажби на услуги, стоки и продукция',[PaymentTypeCode] = '444000', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 10;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от наеми на имущество',[PaymentTypeCode] = '441400', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 11;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от наеми на земя',[PaymentTypeCode] = '444100', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 12;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от лихви по текущи банкови сметки',[PaymentTypeCode] = '444300', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 13;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от лихви по срочни депозити',[PaymentTypeCode] = '443400', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 14;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Дивидент',[PaymentTypeCode] = '444800', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 15;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Конфискувани средства и приходи от продажби на конфискувани вещи',[PaymentTypeCode] = '444900', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 16;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Дарения,помощи и други безвъзмездно получени суми от страната',[PaymentTypeCode] = '445100', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 17;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Дарения,помощи и други безвъзмездно получени суми от чужбина',[PaymentTypeCode] = '445200', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 18;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от продажби на дълготрайни материални активи',[PaymentTypeCode] = '445500', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 19;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от продажби на земя',[PaymentTypeCode] = '445600', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 20;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от концесии',[PaymentTypeCode] = '445700', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 21;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от ликвидиране на общински предприятия',[PaymentTypeCode] = '445800', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 22;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от продажба на нематериални активи',[PaymentTypeCode] = '445900', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 23;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Глоби, санкции и неустойки,нак. лихви, обезщетения и начети',[PaymentTypeCode] = '446500', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 24;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Други неданъчни приходи',[PaymentTypeCode] = '447000', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 25;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за технически услуги',[PaymentTypeCode] = '448001', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 26;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на детски градини',[PaymentTypeCode] = '448002', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 27;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на детски ясли и други по здравеопазването',[PaymentTypeCode] = '448003', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 28;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на детски лагери и други по социалния отдих',[PaymentTypeCode] = '448004', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 29;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на домашен социален патронаж и други социални услуги',[PaymentTypeCode] = '448005', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 30;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Окончателен годишен (патентен) данък и данък върху таксиметров превоз на пътници',[PaymentTypeCode] = '441400', [AlgorithmId] = 0, [IsActive] = 1 WHERE [ObligationTypeId] = 31;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на пазари, тържища и др.',[PaymentTypeCode] = '448008', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 32;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на полудневни детски градини',[PaymentTypeCode] = '448009', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 33;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за ползване на общежития и други по образованието',[PaymentTypeCode] = '448010', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 34;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за гробни места',[PaymentTypeCode] = '448011', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 35;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Такси за притежаване на куче',[PaymentTypeCode] = '448013', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 36;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Други общински такси',[PaymentTypeCode] = '448090', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 37;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Приходи от такси за общините по Закона за опазване на околната среда',[PaymentTypeCode] = '448081', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 38;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Неразпределени общински приходи за операции по чл. 4, ал. 1 от ЗОПБ',[PaymentTypeCode] = '448844', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 39;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Неидентифицирани преводи по сметки 7311 (т. 21 от ДДС № 06/2006 г. )',[PaymentTypeCode] = '448868', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 40;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'Преводи по с/ки 8 4 с многоредов документ',[PaymentTypeCode] = '448888', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 41;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'дебитно - контролираща сметка - код за централизация към сметка 7304',[PaymentTypeCode] = '449944', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 42;
UPDATE  [dbo].[ObligationTypes] SET [Name] = 'дебитно - контролираща сметка - код за преводи от сметка 7304',[PaymentTypeCode] = '449998', [AlgorithmId] = 1, [IsActive] = 1 WHERE [ObligationTypeId] = 43;


UPDATE [dbo].[GlobalValues]
	SET Value= '30'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2022-04-14_Version30', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2022-04-14_Version30 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO
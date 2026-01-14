BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '4';
DECLARE @ScriptName NVARCHAR(MAX) = '2015-12-10_Version4';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

SET IDENTITY_INSERT [dbo].[EserviceClients] ON 

INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint])
VALUES (4, 1, N'ArchimedLovech', N'archimed_eProcess_Lovech_e39460e1-ce74-48f9-845f-3703d0781b64', N'ИС "Archimed eProcess", гр. Ловеч', N'', NULL, 'Банка ДСК', 'STSABGSF', '', N'SKLDJFHLERTJKSJH76WTUY23894690XM345SDaMCERNJ4WEYUCCVNNKPWEUXCNV5MBNSRKJDTHEKJFDVHUKHEJ276489CVVZGFWURTOTONXCVMNSRUH569034NNXXMNC', 1, N'', N'', '7e74b8b0-9f5a-4889-86a8-ee522121021a', 0, 0)

INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint])
VALUES (5, 1, N'ArchimedStaraZagora', N'archimed_eProcess_StaraZagora_47adf7ba-2843-4097-8208-6aed116d1f55', N'ИС "Archimed eProcess", гр. Стара Загора', N'', NULL, 'Банка ДСК', 'STSABGSF', '', N'ASD567DFGJH54K62323UYSDFVNXCVTYOPOIWYGHSKLYUIBNCFGTI27HHHCVBASDSFIKU32ZMKLPIASFRQSDHGZXCBHSJKHG234BNKBVXKJHFKDJHJ234628CBSHD52CV', 1, N'', N'', 'a4f6001a-5465-44e1-be72-1847c4b386fc', 0, 0)

INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint])
VALUES (6, 1, N'ArchimedHaskovo', N'archimed_eProcess_Haskovo_3add2798-8bde-4ac9-bc55-eeb46e6e5a63', N'ИС "Archimed eProcess", гр. Хасково', N'', NULL, 'Банка ДСК', 'STSABGSF', '', N'YYTURTGCVBDY467DFNZXCNWEIU2UHDFVXCMBSKDFG34MNXBCSGERKERHERUKTH958ZGHASFDLSDHKLPETIYJHXCXBVGTOXCHSDFJ234456MXCVHNKSFGIPORUZNBNFU3', 1, N'', N'', 'f3702984-8b1d-42ea-9249-7865d5c47872', 0, 0)

SET IDENTITY_INSERT [dbo].[EserviceClients] OFF


--------------------------------------------------------------
--Update database version
Update [dbo].[GlobalValues] 
SET [Value] = @DbVersion,
    [ModifyDate] = GETDATE()
WHERE [Key]='DatabaseVersion'
--Add DatabaseUpdate record
INSERT INTO [dbo].[DatabaseUpdates] (ScriptName, Notes, UpdateDate)
VALUES (@ScriptName, @Notes, GETDATE())
-------------------------------------------------------------
--End commands block
COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH







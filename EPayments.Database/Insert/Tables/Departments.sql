PRINT 'Insert [Departments]'
GO

SET IDENTITY_INSERT [dbo].[Departments] ON 
GO

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (1, N'Министерство на правосъдието')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (2, N'Националният център за информация и документация (НАЦИД)')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (3, N'Изпълнителна агенция „Автомобилна администрация“')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (4, N'Национален осигурителен институт (НОИ)')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (5, N'Национален статистически институт')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (6, N'Български Институт по Метрология (БИМ)')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (7, N'Община Радомир')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (8, N'Община Габрово')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (9, N'Община Бургас')

INSERT [dbo].[Departments] ([DepartmentId], [Name])
VALUES (99, N'Тестово ведомство')

SET IDENTITY_INSERT [dbo].[Departments] OFF
GO

CREATE TABLE [dbo].[EserviceAdminUsers](
	[EserviceAdminUserId] [int] IDENTITY(1,1) NOT NULL,
	[DepartmentId] [int] NULL,
	[Username] [nvarchar](100) NOT NULL,
	[PasswordHash] [nvarchar](max) NOT NULL,
	[PasswordSalt] [nvarchar](max) NOT NULL,
	[IpAddressesForAccess] [nvarchar](max) NULL,
	[Name] [nvarchar](100) NOT NULL,
	[Email] [nvarchar](100) NULL,
	[InsufficientAmountNotifications] [bit] NOT NULL,
	[OverpaidAmountNotifications] [bit] NOT NULL,
	[ReferencedNotifications] [bit] NOT NULL,
	[NotReferencedNotifications] [bit] NOT NULL,
	[IsActive] [bit] NOT NULL,
	[ReferringEserviceClientId] [int] NULL,
 CONSTRAINT [PK_EserviceAdminUsers] PRIMARY KEY CLUSTERED 
(
	[EserviceAdminUserId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Username] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceAdminUsers]  WITH CHECK ADD FOREIGN KEY([ReferringEserviceClientId])
REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
GO
ALTER TABLE [dbo].[EserviceAdminUsers]  WITH CHECK ADD  CONSTRAINT [FK_EserviceAdminUsers_Departments] FOREIGN KEY([DepartmentId])
REFERENCES [dbo].[Departments] ([DepartmentId])
GO

ALTER TABLE [dbo].[EserviceAdminUsers] CHECK CONSTRAINT [FK_EserviceAdminUsers_Departments]
GO
ALTER TABLE [dbo].[EserviceAdminUsers]  WITH CHECK ADD  CONSTRAINT [FK_EserviceAdminUsers_EserviceAdminUsers] FOREIGN KEY([EserviceAdminUserId])
REFERENCES [dbo].[EserviceAdminUsers] ([EserviceAdminUserId])
GO

ALTER TABLE [dbo].[EserviceAdminUsers] CHECK CONSTRAINT [FK_EserviceAdminUsers_EserviceAdminUsers]
CREATE TABLE [dbo].[DistributionErrors](
	[DistributionErrorId] [int] IDENTITY(1,1) NOT NULL,
	[Error] [nvarchar](500) NOT NULL,
	[CreatedAt] [datetime2](7) NOT NULL,
	[DistributionRevenueId] [int] NOT NULL,
 CONSTRAINT [PK_DistributionErrorId] PRIMARY KEY CLUSTERED 
(
	[DistributionErrorId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DistributionErrors]  WITH CHECK ADD  CONSTRAINT [FK_DistributionErrors_DistributionRevenues] FOREIGN KEY([DistributionRevenueId])
REFERENCES [dbo].[DistributionRevenues] ([DistributionRevenueId])
GO

ALTER TABLE [dbo].[DistributionErrors] CHECK CONSTRAINT [FK_DistributionErrors_DistributionRevenues]
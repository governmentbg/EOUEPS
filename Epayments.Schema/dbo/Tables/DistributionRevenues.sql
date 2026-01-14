CREATE TABLE [dbo].[DistributionRevenues](
	[DistributionRevenueId] [int] IDENTITY(1,1) NOT NULL,
	[CreatedAt] [datetime2](7) NOT NULL,
	[DistributedDate] [datetime2](7) NULL,
	[IsDistributed] [bit] NOT NULL,
	[TotalSum] [decimal](18, 4) NOT NULL,
	[DistributionTypeId] [int] NOT NULL,
	[IsFileGenerated] [bit] NOT NULL,
	[FileName] [nvarchar](50) NULL,
 CONSTRAINT [PK_DistributionRevenues] PRIMARY KEY CLUSTERED 
(
	[DistributionRevenueId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DistributionRevenues]  WITH CHECK ADD  CONSTRAINT [FK_DistributionRevenues_DistributionTypes] FOREIGN KEY([DistributionTypeId])
REFERENCES [dbo].[DistributionTypes] ([DistributionTypeId])
GO

ALTER TABLE [dbo].[DistributionRevenues] CHECK CONSTRAINT [FK_DistributionRevenues_DistributionTypes]
GO
ALTER TABLE [dbo].[DistributionRevenues] ADD  DEFAULT ((0)) FOR [IsDistributed]
GO
ALTER TABLE [dbo].[DistributionRevenues] ADD  DEFAULT ((0)) FOR [IsFileGenerated]
USE [Northwind]
GO

/****** Object:  StoredProcedure [dbo].[spInsertProducts]    Script Date: 18/10/2016 19:51:08 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spInsertProducts](
	@ProductName nvarchar(50),
	@SupplierId int,
    @UnitPrice decimal(12,2),
    @Package nvarchar(30),
    @IsDiscontinued bit = 0)
AS
BEGIN
INSERT INTO Product(
	ProductName, 
	SupplierId, 
	UnitPrice,
	Package,
	IsDiscontinued
	)
VALUES (
	@ProductName,
	@SupplierId,
	@UnitPrice,
	@Package,
	@IsDiscontinued
)
END

GO


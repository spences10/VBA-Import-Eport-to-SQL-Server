USE [Northwind]
GO

/****** Object:  StoredProcedure [dbo].[spUpdateProducts]    Script Date: 18/10/2016 19:51:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[spUpdateProducts]
    @Id int,
	@ProductName nvarchar(50),
	@SupplierId int,
    @UnitPrice decimal(12,2),
    @Package nvarchar(30),
    @IsDiscontinued bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE Product
    SET 
		ProductName=@ProductName, 
		SupplierId=@SupplierId, 
		UnitPrice=@UnitPrice,
		Package=@Package,
		IsDiscontinued=@IsDiscontinued
    WHERE Id=@id
END

GO


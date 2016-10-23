USE [Northwind]
GO

/****** Object:  StoredProcedure [dbo].[spGetAllSupplierProducts]    Script Date: 18/10/2016 19:50:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spGetAllSupplierProducts]
    @SupplierId int
AS   
    SET NOCOUNT ON;  
	SELECT
	P.Id,
	S.CompanyName 'Company Name',
	P.ProductName 'Product Name',
	P.UnitPrice 'Unit Price',
	P.Package
FROM 
	Supplier S
	LEFT JOIN Product P 
	ON S.ID=P.SupplierId
WHERE
	P.IsDiscontinued = 0 AND 
	P.SupplierId=@SupplierId
ORDER BY
	P.ProductName

GO


# VBA Import Export to SQL server


# Description

This is an example of using Excel for a front end for SQL server using Excel to validate your data.	

It has validation in it to notify the user of the database constraints.

# Build

This whole project is solely dependant on you having either a local instance SQL server or a SQL server you can create the sample-model.sql database in.

The files sample-model.slq and sample-data.sql should be used to create the database and data for this project, the scripts were taken from http://www.dofactory.com/sql/sample-database which is a simple database for this project to run from.

Once the database and data scripts have been run then you will need to create the stored procs used by the project spGetAllSupplierProducts.sql, spInsertProducts.sql, spSupplierList.sql and spUpdateProducts.sql.

ImportExport.xlsm is stored as a BLOB instead of being created say via a ps script as it has worksheet controls and some code contained for worksheet events.

You will need to import modImportExport.bas and modValidation.bas, you should then be able to run the Import from the Control sheet of the ImportExport.xlsm workbook.

# Still to come

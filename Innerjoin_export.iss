Sub Main
	Call JoinDatabase()	'Sample-Detailed Sales.IMD
	Call ExportDatabaseXLSX()	'Join Databases.IMD
End Sub


' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Sample-Detailed Sales.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Sample-Detailed Sales Previous Year.IMD"
	task.AddPFieldToInc "SALESREP_NO"
	task.AddPFieldToInc "CUSTNO"
	task.AddPFieldToInc "UNIT_PRICE"
	task.AddPFieldToInc "QTY"
	task.AddPFieldToInc "SALES_BEF_TAX"
	task.AddPFieldToInc "SALES_TAX"
	task.AddPFieldToInc "SALES_PLUS_TAX"
	task.AddSFieldToInc "SALESREP_NO"
	task.AddSFieldToInc "CUSTNO"
	task.AddSFieldToInc "UNIT_PRICE"
	task.AddSFieldToInc "QTY"
	task.AddSFieldToInc "SALES_BEF_TAX"
	task.AddSFieldToInc "SALES_TAX"
	task.AddSFieldToInc "SALES_PLUS_TAX"
	task.AddMatchKey "SALESREP_NO", "CUSTNO", "A"
	task.CreateVirtualDatabase = False
	dbName = "Join Databases.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_REC
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Join Databases.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = "SALES_BEF_TAX > 5000"
	task.PerformTask "C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\Join Databases.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
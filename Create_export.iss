Sub Main
	Call DirectExtraction()	'Sample-Bank Transactions.IMD
	Call ExportDatabaseXLSX()	'High Amount.IMD
End Sub


' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Sample-Bank Transactions.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "High Amount.IMD"
	task.AddExtraction dbName, "", "Amount > 5000"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("High Amount.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\High Amount.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
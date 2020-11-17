Dim dbName as String 

Sub Main
	Call ExportDatabase()
End Sub


Function ExportDatabase()
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Index
	task.AddKey "NAME", "A"
	task.Index FALSE
	task = db.ExportDatabase
	task.IncludeAllFields
	' Display the setup dialog box before performing the task.
	task.DisplaySetupDialog 0
	Set db = Nothing
	Set task = Nothing
End Function
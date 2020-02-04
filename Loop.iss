Dim Num As Integer

Sub Main

	Call NumberOfPulls()
	Count = 0
	Do While Count < Num
		Call ExcelImport()
	Count = Count + 1
Loop


	

End Sub


Function NumberOfPulls
	subFileName = InputBox("How many sheets you want to pull: ", "Name Input", "1")
	Num  = Val(subFileName)
End Function 

' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		dbName =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(dbName ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
End Function


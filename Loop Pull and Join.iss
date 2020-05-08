Dim importedFile As String
Dim Num As Integer
Dim CleanYearDatabase(50) As String
Dim PrimeDatabase As String
Dim SecondDatabase As String
Dim NewDatabaseName As String
Dim singleDatabase As String 
Dim tempFileName As String 


Sub Main
	Call NumberOfPulls() 
	i = 0
	Do While i < Num
		Call ExcelImport(i)
		Call CleanYear()
		i = i +1
		Client.RefreshFileExplorer
	Loop
	'If there is only one database we will not need to do a join. If there is then it will run through the join script 
	If Num > 1 Then 
		j = 0
		Do While j + 1 < Num
			Call DatabaseToJoin(j)
			Call JoinDatabase(PrimeDatabase)
			j = j + 1
			Client.RefreshFileExplorer
		Loop
	ElseIf Num = 1 Then

	End If
End Sub

Function NumberOfPulls
	subFileName = InputBox("How many sheets you want to pull: ", "Name Input", "1")
	Num  = Val(subFileName)
End Function 

Function DatabaseToJoin(j)
	If j = 0 Then 
		PrimeDatabase = InputBox("Enter primary database: ", "Name Input", "[Year][Month]TransactionStatement.xlsx Clean")
		PrimeDatabase = PrimeDatabase + ".IMD"
	ElseIf j > 0 Then 
		PrimeDatabase = InputBox("Enter primary database: ", "Database")
	End If
	SecondDatabase = InputBox("Enter secondary database: ", "Name Input", "[Year][Month]TransactionStatement.xlsx Clean")
	SecondDatabase = SecondDatabase + ".IMD"
	NewDatabaseName = InputBox("Enter the neam of the new database: ", "Name Input", "Database")
End Function

' File - Import Assistant: Excel
Function ExcelImport(i)
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		importedFile =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = importedFile
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(importedFile ,"","\",1,1)
	importedFile =  iSplit(importedFile ,"","\",1,1)
	tempFileName = importedFile
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	importedFile = task.OutputFilePath("Sheet1")
	'adding the name of the new database into the array
	CleanYearDatabase(i) = importedFile 
	Set task = Nothing
End Function

' Data: Direct Extraction
Function CleanYear
	Set db = Client.OpenDatabase(importedFile)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "SHORT_NAME"
	task.AddFieldToInc "ACCOUNT_NUMBER"
	task.AddFieldToInc "MANAGING_ACCOUNT_NUMBER"
	task.AddFieldToInc "MANAGING_ACCOUNT_NAME"
	task.AddFieldToInc "MANAGING_ACCOUNT_NAME_LINE_2"
	task.AddFieldToInc "SOCIAL_SECURITY_NUMBER"
	task.AddFieldToInc "OPTIONAL_1"
	task.AddFieldToInc "OPTIONAL_2"
	task.AddFieldToInc "CURRENT_DEFAULT_ACCOUNTING_CODE"
	task.AddFieldToInc "LOST_STOLEN_ACCOUNT"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "POSTING_DATE"
	task.AddFieldToInc "CYCLE_CLOSE_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "SOURCE_CURRENCY_AMOUNT"
	task.AddFieldToInc "SOURCE_CURRENCY"
	task.AddFieldToInc "SALES_TAX"
	task.AddFieldToInc "POSTING_TYPE"
	task.AddFieldToInc "PURCHASE_ID"
	task.AddFieldToInc "TRANSACTION_STATUS"
	task.AddFieldToInc "DISPUTED_STATUS"
	task.AddFieldToInc "DISPUTE_STATUS_DATE"
	task.AddFieldToInc "REFERENCE_NUMBER"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_NAME"
	task.AddFieldToInc "MERCHANT_CITY"
	task.AddFieldToInc "MERCHANT_STATE_PROVINCE"
	task.AddFieldToInc "TAXPAYER_ID_NUMBER_TIN"
	task.AddFieldToInc "MERCHANT_ORDER_NUMBER"
	task.AddFieldToInc "MEMO_TO_ACCOUNT_NAME"
	task.AddFieldToInc "MEMO_TO_ACCOUNT_NUMBER"
	task.AddFieldToInc "POSTED_TO_ACCOUNT_NAME"
	task.AddFieldToInc "POSTED_TO_ACCOUNT_NUMBER"
	task.AddFieldToInc "BILLING_TYPE"
	task.AddFieldToInc "CLIENT_NAME"
	task.AddFieldToInc "REPORT_DATE"
	task.AddFieldToInc "REPORT_NAME"
	task.AddFieldToInc "DATE_TYPE"
	task.AddFieldToInc "START_DATE"
	task.AddFieldToInc "END_DATE"
	task.AddFieldToInc "REVIEWED_STATUS"
	task.AddFieldToInc "DISPUTED_STATUS1"
	task.AddFieldToInc "TRANSACTION_AMOUNT1"
	task.AddFieldToInc "POSTING_TYPE1"
	task.AddFieldToInc "ALLOCATION_DETAIL"
	task.AddFieldToInc "TRANSACTION_COMMENTS"
	task.AddFieldToInc "TRANSACTION_CUSTOM_FIELDS"
	task.AddFieldToInc "FLEET_DETAIL"
	task.AddFieldToInc "PAYMENTS"
	task.AddFieldToInc "FEES"
	task.AddFieldToInc "INCLUDE_PROCESSING_HIERARCHY_NAMES"
	task.AddFieldToInc "SORT_1"
	task.AddFieldToInc "SORT_2"
	task.AddFieldToInc "SORT_3"
	task.AddFieldToInc "SORT_4"
	task.AddFieldToInc "BANK"
	task.AddFieldToInc "AGENT"
	task.AddFieldToInc "COMPANY"
	task.AddFieldToInc "DIVISION"
	task.AddFieldToInc "DEPARTMENT"
	singleDatabase = tempFileName + " Clean.IMD"
	task.AddExtraction singleDatabase, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase importedFile
	Set task = Nothing
	Set db = Nothing
End Function

' File: Join Databases
Function JoinDatabase(PrimeDatabase)
	Set db = Client.OpenDatabase(PrimeDatabase)
	Set task = db.AppendDatabase
	task.AddDatabase SecondDatabase
	dbName = NewDatabaseName
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function



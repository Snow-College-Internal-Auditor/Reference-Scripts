'test
Sub Main

	Call createFolder()
	Call moveDatabase()

End Sub 

Function createFolder

	' Set the task type.
	Set task = Client.ProjectManagement
	
	' Create a new folder.
	task.CreateFolder "IDEATest"
	Set task = Nothing

End Function

Function moveDatabase
	
	' Declare variables and objects.
	Dim path As String
	Dim pm As Object
	
	' Access project management object to manage databases/projects on
	' server.
	Set pm = Client.ProjectManagement
	
	' Use path object to get the full path and file name to the specified database.
	Set path = "C:\Users\mckinnin.lloyd\Documents\My IDEA Documents\IDEA Projects\Samples\BEAUTY.IMD"
	
	' Move the file from the server to a different server location.
	pm.MoveDatabase path, "IDEATest"
	
	' Refresh the File Explorer.
	Client.RefreshFileExplorer
	
	' Clear the path.
	Set pm = Nothing
	
End Function

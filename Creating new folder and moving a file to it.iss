Dim subFilename As String

Sub Main

	Call createFolder()
	Call moveDatabase()

End Sub 

Function createFolder

	' Set the task type.
	Set task = Client.ProjectManagement
	
	subFilename = InputBox("Type The Name of The Month: ", "Name Input", "IDEATest")
	
	' Create a new folder.
	task.CreateFolder subFilename
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
	Set path = "VIDEO"
	
	' Move the file from the server to a different server location.
	pm.MoveDatabase path, subFilename
	
	' Refresh the File Explorer.
	Client.RefreshFileExplorer
	
	' Clear the path.
	Set pm = Nothing
	
End Function

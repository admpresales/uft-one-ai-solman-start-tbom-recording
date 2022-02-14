'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Script limitations:
'		The script doesn't handle if the TCODE is in multiple BPH nodes, it will only create the TBOM work list item for the first BPH node in the list.
'			This is due to a limitation of SAP Solution Manager, when you ask it to create multiple work list items, it only tells you the task number
'			for the first record that you selected.
'		The script doesn't handle if the status of the TBOM is Obsolete because the dev system didn't have any Obsolete status TBOMs.
'		The script assumes that if there is no TBOM for the TCode, there is no Work Item created either.  If there is, the script will select the originally
'			created workitem for that TCode (due to that same Solution Manager limitation), even though it may create another work item too.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim BrowserExecutable, oShell, counter, fso, filespec

'Deletes the .sap file if the file exists
Set fso = CreateObject("Scripting.FileSystemObject")
filespec = "C:\Users\demo\Downloads\ags_work_appln.sap"
If (fso.FileExists(filespec)) Then
	fso.DeleteFile(filespec)
End If
Set fso = Nothing

'Starts the mediaserver service (even if it is already started).
Set oShell = CreateObject ("WSCript.shell")
oShell.run "powershell -command ""Start-Service mediaserver"""
Set oShell = Nothing

While Browser("CreationTime:=0").Exist(0)   													'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3												'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")												'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")												'Navigate to the application URL
AppContext.Maximize																		'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																			'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

AIUtil("text_box", micAnyText, micWithAnchorOnLeft, AIUtil.FindText("User")).Type DataTable.Value("Login")
AIUtil("text_box", "Password").Type DataTable.Value("Password")
AIUtil("button", "Log On").Click

counter = 0
Do
	counter = counter + 1
	wait 1
	If counter >= 60 Then
		'msgbox "The search icon didn't show up within " & counter & " tries, check application."
		Reporter.ReportEvent micFail, "Login", "Login failed.  The search icon didn't show up within " & counter & " tries, check application."
		ExitTestIteration
	End If
	
	If counter>=3 Then
		If AIUtil("button", "Continue").Exist(0) Then
			'AIUtil("check_box", "Cancel all existing Iogons").SetState "Off"
			AIUtil("button", "Continue").Click
		End If
	End If
	
Loop Until AIUtil("search").Exist(0)


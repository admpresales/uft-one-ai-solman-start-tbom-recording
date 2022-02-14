Set AppContext=Browser("CreationTime:=0")												'Set the variable for what application (in this case the browser) we are acting upon
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

AIUtil.FindText("Process Step Library",micFromLeft,1).Click

AIUtil.FindTextBlock("Solution > Libraries > Process Step Library").CheckExists True
AIUtil("search", micAnyText, micFromLeft, 1).Search DataTable.Value("TCode")
AppContext.Sync																			'Wait for the browser to stop spinning

If AIUtil.FindText("No Elements Selected").Exist Then
	AIUtil("down_triangle", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Filter by")).Click
	AIUtil.FindTextBlock("Type", micFromBottom, 1).Click
ElseIf AIUtil.FindText("Your search did not match").Exist Then
	'msgbox "The TCode isn't associated with any BPH node"
	AppContext.Close
	Reporter.ReportEvent micFail, "Check TCode in BPH", "The TCode " & DataTable.Value("TCode") & " isn't associated with any BPH node.  Please set the executable to a BPH node and re-run"
	ExitTestIteration
End If
If AIUtil.FindTextBlock("Type v").Exist Then
	'Type filter applied
Else
	AppContext.Close
	Reporter.ReportEvent micFail, "Type Filter", "The Type filter didn't apply, check application."
	ExitTestIteration
End If

If AIUtil("check_box", micAnyText, micWithAnchorOnRight, AIUtil.FindText("TBOM",micFromLeft,1)).Exist(0) Then
	AIUtil("check_box", micAnyText, micWithAnchorOnRight, AIUtil.FindText("TBOM",micFromLeft,1)).SetState "On"
	AIUtil.FindTextBlock("Name").Click
	AIUtil("plus").Click
	AIUtil("search", micAnyText, micFromTop, 2).Search "TBOM"
	AIUtil("check_box", "TBOM Status", micWithAnchorBelow, AIUtil.FindTextBlock("TBOM Status")).SetState "On"
	AIUtil.FindTextBlock("OK").Click
	CurrentTBOMStatus = AIUtil.FindTextBlock(micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("TBOM", micFromBottom, 1)).GetText
	If ((CurrentTBOMStatus = "Updated (U)") or (CurrentTBOMStatus = "Created (C)"))Then
		'msgbox "TBOM is current, exit"
		AppContext.Close
		Reporter.ReportEvent micFail, "TBOM Status", "The TCode " & DataTable.Value("TCode") & " has a current TBOM, so no further action is needed, exiting test."
		ExitTestIteration
	End If
	msgbox CurrentTBOMStatus
Else
	'msgbox "No TBOM record exists yet"
	AIUtil("check_box", micAnyText, micWithAnchorOnRight, AIUtil.FindText("Transaction <Exec.Ref.>",micFromLeft,1)).SetState "On"
	counter = 0
	Do
		counter = counter + 1
		AIUtil("check_box", "Name").SetState "On"
		If counter >= 60 Then
			'msgbox "The search icon didn't show up within " & counter & " tries, check application."
			Reporter.ReportEvent micFail, "Setting Check Box On", "The No Elements Selected is still displaying after " & counter & " tries, check application."
			ExitTestIteration
		End If
		wait 1
	Loop While AIUtil.FindText("No Elements Selected").Exist(0)
	Setting.WebPackage("ReplayType") = 2
	Browser("Solution Documentation").Page("Solution Documentation").WebElement("First_Transaction_Exec_Ref").RightClick
	Setting.WebPackage("ReplayType") = 1
	AIUtil.FindTextBlock("Create TBOM Work Items").Click
	Set AppContext=Browser("CreationTime:=1")												'Set the variable for what application (in this case the browser) we are acting upon
	AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
	'AIUtil("check_box", micAnyText, micWithAnchorBelow, AIUtil.FindTextBlock("State", micFromTop, 1)).SetState "On"
	AIUtil("check_box", "Create", micFromTop, 1).SetState "On"
	AIUtil("button", "Apply Mass Change").Click
	AIUtil("text_box", "BP Expert:").Type "302"
	AIUtil("button", "Apply").Click
	AIUtil("button", "Create Work Items").Click
	If AIUtil.FindTextBlock("Work items created. Check their status").Exist Then
		AIUtil.FindTextBlock("Work items created. Check their status").Click
	Else
		AIUtil("close", micAnyText, micFromTop, 1).Hover
		AIUtil.Context.SetBrowserScope(BrowserWindow)
		AIUtil.RunSettings.AutoScroll.Disable
		If AIUtil.FindText("Error Message").Exist Then
			Reporter.ReportEvent micFail, "TBOM Work Item Creation", "The TCode " & DataTable.Value("TCode") & " has errors preventing the creation of the TBOM Work Item.  Please resolve and re-run."
			AIUtil.RunSettings.AutoScroll.Enable "down", 2
			AIUtil.Context.SetBrowserScope(WebPage)
			While Browser("CreationTime:=0").Exist(0)   													'Loop to close all open browsers
				Browser("CreationTime:=0").Close 
			Wend
			ExitTestIteration
		End If
		AIUtil.RunSettings.AutoScroll.Enable "down", 2
		AIUtil.Context.SetBrowserScope(WebPage)
	End If
	WorkItemText = AIUtil.FindTextBlock(micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("State:", micFromBottom, 1)).GetText
	WorkItemArray = split(WorkItemText, " ")
	DataTable.Value("WorkItemNumber") =  WorkItemArray(3)
End If
While Browser("CreationTime:=0").Exist(0)   													'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend

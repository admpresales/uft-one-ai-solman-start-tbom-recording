Set AppContext=Browser("CreationTime:=0")												'Set the variable for what application (in this case the browser) we are acting upon
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

AIUtil.FindText("Process Step Library",micFromLeft,1).Click
AIUtil.FindTextBlock("List").Click
AppContext.Sync																			'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Browser").Click
AppContext.Sync																			'Wait for the browser to stop spinning
AIUtil("search", micAnyText, micFromLeft, 1).Search DataTable.Value("TCode")
AppContext.Sync																			'Wait for the browser to stop spinning

If AIUtil.FindText("No Elements Selected").Exist Then
	AIUtil("down_triangle", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Filter by")).Click
	AIUtil.FindTextBlock("Type", micFromBottom, 1).Click
ElseIf AIUtil.FindText("Your search did not match").Exist Then
	msgbox "The TCode isn't associated with any BPH node"
	AppContext.Close
	ExitTest
End If

'If AIUtil("check_box", "TBOM").Exist Then
If AIUtil("check_box", micAnyText, micWithAnchorOnRight, AIUtil.FindText("TBOM",micFromLeft,1)).Exist Then
	AIUtil("check_box", micAnyText, micWithAnchorOnRight, AIUtil.FindText("TBOM",micFromLeft,1)).SetState "On"
	AIUtil.FindTextBlock("Name").Click
	AIUtil("plus").Click
	AIUtil("search", micAnyText, micFromTop, 2).Search "TBOM"
	AIUtil("check_box", "TBOM Status", micWithAnchorBelow, AIUtil.FindTextBlock("TBOM Status")).SetState "On"
	AIUtil.FindTextBlock("OK").Click
	CurrentTBOMStatus = AIUtil.FindTextBlock(micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("TBOM", micFromBottom, 1)).GetText
	If CurrentTBOMStatus = "Updated (U)" Then
		msgbox "TBOM is current, exit"
		AppContext.Close
		ExitTest
	End If
	msgbox CurrentTBOMStatus
Else
	msgbox "No TBOM record exists yet"
End If


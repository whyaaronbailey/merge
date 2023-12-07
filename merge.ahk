TraySetIcon A_ScriptDir "\Data\merge.ico" ;set the tray icon
Run A_ScriptDir "\Data\powermic_control.exe" ;run the PowerMic hotkey adapter
#InputLevel 1 ;set inputlevel to be able to receive scribe
ProcessSetPriority "High"
#SingleInstance Force

~::
{
	global ;set variables as global
	destroyed := ""
	SetTitleMatchMode 1

;; close open patients in Epic
	if  WinExist("ahk_exe alihmicafhost.exe") {
	  CloseEpic()
	}

;; get accession number from PowerScribe360
	If WinExist("ahk_exe Nuance.PowerScribe360.exe") {
	   acc := GetStudyInfo()
	}

;; send accession number to Epic
	if  WinExist("ahk_exe alihmicafhost.exe")
		SendtoEpic()
	
;; generate GUI
	MyGui := Gui()
	MakeGui()
	MyGui.Add('Text','w200', '')	
	
	;; Place GUI Near GE PACs
	if WinExist("Study List")
	{
		WinGetPos &pacsx, &pacsy, &pacsw, &pacsh, "Study List"
		guix := pacsx + 115
		guiy := pacsy + 240
		guioptions := "w225 x" guix " y" guiy
	} else {
		guioptions := "w225 x100 y500"
	}
	MyGui.Show(guioptions)
	MyGui.Opt("+AlwaysOnTop")
	MyGui.Title := "Merge" 
	SetTimer CloseMerge, 30000 ;; autocloses the GUI window in 30 seconds
	return
}

;; open accession number in GE PACs
EnterInformation(GuiCtrlObj, Info)
{
  If WinExist("Study List") {
	  WinActivate ("Study List")
	  ControlFocus "Internet Explorer_Server1", "Study List"
	  SetKeyDelay(5, 5)
	  ControlSendText(GuiCtrlObj.text, "Internet Explorer_Server1", "Study List")
	  ControlSend("{enter}", "Internet Explorer_Server1", "Study List")
  }
  A_Clipboard := GuiCtrlObj.text
  destroyed := "no"
  return
}

;; open accession number in GE PACs
EnterInformationandKill(GuiCtrlObj, Info)
{ 
  If WinExist("Study List") {
	  WinActivate ("Study List")
	  ControlFocus "Internet Explorer_Server1", "Study List"
	  SetKeyDelay(5, 5)
	  ControlSendText(GuiCtrlObj.text, "Internet Explorer_Server1", "Study List")
	  ControlSend("{enter}", "Internet Explorer_Server1", "Study List")
  }
  A_Clipboard := GuiCtrlObj.text
  MyGui.Destroy()
  destroyed := "yes"
  return
}

SendtoEpic()
{
	SetTitleMatchMode 1
	#WinActivateForce
	WinActivate ("Hyperspace")
	SetControlDelay -1
	ControlClick "Chrome_RenderWidgetHostHWND1", "Hyperspace"
	ControlSend("{control down}2{control up}", "Chrome_RenderWidgetHostHWND1", "Hyperspace")
	Sleep(1750)
	Send("{control down}v")
	Send("{control up}")
	Sleep (500)
	Send ("{enter}")
	Sleep (500)
	Send("{enter}")
	return
}

CloseMerge()
{
  MyGui.Destroy()
  return
}

CloseEpic()
{
	SetTitleMatchMode 1
	#WinActivateForce
	WinActivate ("Hyperspace")
	SetControlDelay -1
	ControlClick "Chrome_RenderWidgetHostHWND1", "Hyperspace"
	ControlSend("{control down}w{control up}", "Chrome_RenderWidgetHostHWND1", "Hyperspace")
	Sleep (500)
	return
}

GetStudyInfo()
{
	#WinActivateForce
	WinActivate("PowerScribe")
	WinWaitActive("ahk_exe Nuance.PowerScribe360.exe")
	controls := WinGetControls("ahk_exe Nuance.PowerScribe360.exe")
	controlText := ""
	for index, control in controls {
		text := ControlGetText(control)
		if (InStr(text, "Report -") = 1 || InStr(text, "Addendum -") = 1) {
			rightControl := control
			controlText := text
			break
		}
	}

	if (controlText != "") {
		controlParts := StrSplit(controlText, " - ")
		acc := StrSplit(controlParts[3], ",")
		for index, value in acc {
			acc[index] := Trim(value)
		}
	} else {
		MsgBox("No accession number was found. Please retry after opening a report in Powerscribe.")
		exit
	}
	
	A_Clipboard := acc[1]
	return acc
}

Update()
{
	thisVersion := FileRead(A_ScriptDir "\Data\version.txt")
	releaseURL := "https://www.dropbox.com/scl/fi/ve2ya140rz5kn029ihqgn/merge.zip?rlkey=sgc31a9f27nunpwn8eby3bv22&dl=0"
	releaseVersionURL := "https://www.dropbox.com/scl/fi/ncudpqjiodzno0frr315w/version.txt?rlkey=epj58ulxafhx9d5tufb5747j1&dl=0"
	whr := ComObject("WinHttp.WinHttpRequest.5.1")
	whr.Open("GET", releaseVersionURL, true)
	whr.Send()
	whr.WaitForResponse()
	releaseVersion := whr.ResponseText
	if thisVersion != releaseVersion {
		msgBox "Current version: " thisVersion "  Release version: " releaseVersion ". Press OK to update"
		Download releaseURL, A_ScriptDir "\merge.zip"
		run A_ScriptDir "\Data\unzip.exe -o ..\merge.zip"
		MsgBox "Press OK to restart Merge."
		run A_ScriptDir "\merge.exe"
		exitApp
	}
Return
}

MakeGui()
{
	if (acc.Length > 1) {
			MyGui.Add('Text','w200 x0 y0', '')
			MyGui.SetFont("Bold")
			MyGui.Add('GroupBox', 'w210 x6 y10 r7 Wrap', 'Merge - Multiple Accessions')
			MyGui.SetFont()
			MyGui.Add('Text','w200 x12 y35', '1. In PACs, select')
			MyGui.SetFont('s6')
			MyGui.AddDropDownList("w106 x100 y32 Choose1 Disabled",["Accession Number"])
			MyGui.SetFont()
			MyGui.Add('Text','w200 x12 y60', '2. In PACs, click the input field.')
			MyGui.AddEdit('v1 w40 y55 x165 Disabled',' ')
			MyGui.Add('Text','w200 x12 y85', '3. Click button 1, load study.')
			MyGui.SetFont('s6')
			MyGui.Add("Button", "y82 w60 h20 x145 Disabled", acc[1])
			MyGui.SetFont()
			MyGui.Add('Text','w200 x12 y110', '4. Click button 2, load study.')
			MyGui.SetFont('s6')
			MyGui.Add("Button", "y108 w60 h20 x145 Disabled", acc[2])
			MyGui.SetFont("s7 italic")
			MyGui.Add('Text','x15 y137', 'Switch between studies from the taskbar')
			MyGui.SetFont("")
		for index, accNum in acc {
			button := MyGui.Add("Button", "w200 h30 x12 yp+35", accNum)
			if (index = acc.Length) {
				button.OnEvent("Click", EnterInformationandKill)
			} else {
				button.OnEvent("Click", EnterInformation)
			}
		}
	} else {
			MyGui.Add('Text','w200 x0 y0', '')
			MyGui.SetFont("Bold")
			MyGui.Add('GroupBox', 'w210 x6 y10 r4.5 Wrap', 'Merge - Integrator')
			MyGui.SetFont()
			MyGui.Add('Text','w200 x12 y35', '1. In PACs, select')
			MyGui.SetFont('s6')
			MyGui.AddDropDownList("w106 x100 y32 Choose1 Disabled",["Accession Number"])
			MyGui.SetFont()
			MyGui.Add('Text','w200 x12 y60', '2. In PACs, click the input field.')
			MyGui.AddEdit('v1 w40 y55 x165 Disabled',' ')
			MyGui.Add('Text','w200 x12 y85', '3. Click button + load study.')
			MyGui.SetFont('s6')
			MyGui.Add("Button", "y82 w60 h20 x145 Disabled", acc[1])
			MyGui.SetFont()
			button := MyGui.Add("Button", "w200 h30 x12 y130", acc[1])
			button.OnEvent("Click", EnterInformationandKill)
	}
Return
}
	
TraySetIcon "merge.ico" ;set the tray icon
Run 'powermic_control.exe' ;run the PowerMic hotkey adapter
#InputLevel 1 ;set inputlevel to be able to receive scribe
ProcessSetPriority "High"
#SingleInstance Force
~::
{
	global ;set variables as global
	MyGui := Gui()
	destroyed := ""
	SetTitleMatchMode 1
	#WinActivateForce
	WinActivate("PowerScribe")
	WinWaitActive("ahk_exe Nuance.PowerScribe360.exe")

	
;; get accession number from PowerScribe360
;; reading from the report status bar

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

;; send accession number to Epic
	if  WinExist("ahk_exe alihmicafhost.exe")
		SendtoEpic()
	
;; generate GUI

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
	MyGui.Add('Text','w200', '')
	WinGetPos &pacsx, &pacsy, &pacsw, &pacsh, "Study List"
	guix := pacsx + 115
	guiy := pacsy + 240
	guioptions := "w225 x" guix " y" guiy
	MyGui.Show(guioptions)
	MyGui.Opt("+AlwaysOnTop")
	MyGui.Title := "Merge" 
	if (destroyed = "no") 
		SetTimer CloseMerge, 25000
	return
}

;; open accession number in GE PACs
EnterInformation(GuiCtrlObj, Info)
{
  WinActivate ("Study List")
  ControlFocus "Internet Explorer_Server1", "Study List"
  SetKeyDelay(5, 5)
  ControlSendText(GuiCtrlObj.text, "Internet Explorer_Server1", "Study List")
  ControlSend("{enter}", "Internet Explorer_Server1", "Study List")
  A_Clipboard := GuiCtrlObj.text
  destroyed := "no"
  return
}

;; open accession number in GE PACs
EnterInformationandKill(GuiCtrlObj, Info)
{ global
  WinActivate ("Study List")
  ControlFocus "Internet Explorer_Server1", "Study List"
  SetKeyDelay(5, 5)
  ControlSendText(GuiCtrlObj.text, "Internet Explorer_Server1", "Study List")
  ControlSend("{enter}", "Internet Explorer_Server1", "Study List")
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
;	Sleep(100)
;	ControlSend("{control down}w{control up}", "Chrome_RenderWidgetHostHWND1", "Hyperspace")
;	Sleep(200)
	ControlSend("{control down}2{control up}", "Chrome_RenderWidgetHostHWND1", "Hyperspace")
	Sleep(1750)
	Send("{control down}v")
	Send("{control up}")
	Sleep (250)
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


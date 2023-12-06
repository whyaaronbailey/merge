#include AHKHID.ahk
Menu, Tray, Icon, powermic.ico
#SingleInstance Force

Gui, +LastFound
GuiH := WinExist()
;SendMode, Input

;set input level to be able to activate hotkey of merge script
#InputLevel 0 

;Intercept WM_INPUT messages
WM_INPUT := 0x00FF
OnMessage(WM_INPUT, "InputMsg")

AHKHID_Register(1, 0,GuiH, RIDEV_INPUTSINK + RIDEV_PAGEONLY) ;register device

InputMsg(wParam, lParam) 
{
	Local devh, key
	Critical    ;or otherwise you could get ERROR_INVALID_HANDLE

	;get handle of device
	devh := AHKHID_GetInputInfo(lParam, II_DEVHANDLE)

	If (devh <> -1)
        And (AHKHID_GetDevInfo(devh, DI_DEVTYPE, True) = RIM_TYPEHID)
        And (AHKHID_GetDevInfo(devh, DI_HID_VENDORID, True) = 1364) 
        And (AHKHID_GetDevInfo(devh, DI_HID_PRODUCTID, True) = 4097)
		; device matches Nuance Powermic III
		
    {

		;get the keycode
		key := AHKHID_GetInputInfo(lParam, II_MSE_RAWBUTTONS)  
		if (key == 67108864) ;if the left mouse on the Powermic is clicked
				{
				SendLevel 2 ;set the SendLevel command so that the other script receives input
				SendEvent ~ ;send the hotkey to the other script
				;SendEvent {Control up}
			}
	} 
}
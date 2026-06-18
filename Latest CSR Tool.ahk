#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%


; ==================== THIS CODE WILL NOW RUN ====================

SendMode, Input
SetBatchLines, -1
CoordMode,Pixel,Window
CoordMode,Mouse,Window
CoordMode,ToolTip,Screen
global version:=4.1
global PickingMode := "SPV"
global NSN := {}
global LPColumnIndex
global BatchColumnIndex
xl := ComObjActive("Excel.Application")
messages := []

; Creates Loading Screen ========================================================================================================================================================================
w := 525
h := 300

; Bottom-right
margin := 20
x := A_ScreenWidth - w - margin
y := A_ScreenHeight - h - margin

Gui,Loading: +AlwaysOnTop -Caption +ToolWindow +LastFound
Gui,Loading: Margin, 0, 0

; ==========================
; BACKGROUND IMAGE
; ==========================
Gui,Loading: Add, Picture, x0 y0 w%w% h%h%, % "HBITMAP:*" Create_background_jpeg()

; ==========================
; TITLE PANEL
; ==========================
Gui,Loading: Add, Progress, x110 y130 w305 h80 Disabled Background333333

Gui,Loading: Font, s18 Bold cFFFFFF, Segoe UI
Gui,Loading: Add, Text, x0 y145 w%w% Center BackgroundTrans, LOADING

Gui,Loading: Font, s9 cDDDDDD
Gui,Loading: Add, Text, vStatus x0 y180 w%w% Center BackgroundTrans, Initializing...

; ==========================
; PROGRESS BAR
; ==========================
Gui,Loading: Add, Progress, vBar x110 y215 w305 h14 c4CC2FF Background333333, 0

; ==========================
; PERCENT PANEL
; ==========================
Gui,Loading: Add, Progress, x220 y245 w85 h28 Disabled Background333333

Gui,Loading: Font, s10 Bold cFFFFFF
Gui,Loading: Add, Text, vPercent x220 y250 w85 Center BackgroundTrans, 0`%

; Show bottom-right
Gui,loading: Show, x%x% y%y% w%w% h%h%

; Rounded corners
WinSet, Region, 0-18 w%w% h%h% R18-18

; Fade in
Loop 20
{
    alpha := Min(A_Index * 12, 255)
    WinSet, Transparent, %alpha%
    Sleep 20
}

global loadval:=1
loop,10 {
loadval := Round(loadval, 2)
setload(loadval,"Checking for an update...")
loadval++
;sleep,0
}

sleep,100

global LatestVersion := ""

UrlDownloadToFile, https://raw.githubusercontent.com/ILoveYouBev/workplaceUI/refs/heads/main/version.txt, %A_Temp%\version.txt

FileRead, LatestVersion, %A_Temp%\version.txt
FileDelete, %A_Temp%\version.txt

if (LatestVersion != version)
{
	loop,5 {
	loadval := Round(loadval, 2)
	setload(loadval,"Checking for an update...")
	loadval++
	;sleep,0
	}

    Gui, Loading: -AlwaysOnTop
    MsgBox, 4, Update Available, A new update is available!`n`nCurrent: %version%`nLatest: %LatestVersion%`n`nDo you want to update?

    IfMsgBox, Yes
    {

	FileDelete,%A_ScriptDir%/CSR Tool %latestversion%.ahk
	UrlDownloadToFile, https://github.com/ILoveYouBev/workplaceUI/raw/refs/heads/main/Latest`%20CSR`%20Tool.ahk, %A_ScriptDir%/CSR Tool %latestversion%.ahk

	MsgBox, 64, Update, Updated successfully!
	Run,%A_ScriptDir%/CSR Tool %latestversion%.ahk
	ExitApp
    }
    else
    {
        MsgBox, You chose not to update.
    }
}


Width := 400
Height := 250
a:=1,b:=2,c:=3,d:=4,e:=5,f:=6,g:=7,h:=8,i:=9,j:=10,k:=11,l:=12,m:=13,n:=14,o:=15,p:=16,q:=17,r:=18,s:=19,t:=20,u:=21,v:=22,w:=23,x:=24,y:=25,z:=26
LpNumCount:=1
LPStart:=a
CurrentLetter:=chr(LPStart + 96)
CurrentLP=%A_DD%%A_MM%%A_YYYY%%LPNumCount%%CurrentLetter%



Gui,Main: +LastFound
Gui,Main: Color, 808080
Gui,Main: Font, s12 cD0D0D0 Bold
Gui,Main: Add, Progress, % "x-1 y-1 w" (Width+2) " h31 Background404040 Disabled hwndHPROG"
Gui,Main: Add, Text, % "x0 y0 w" Width " h30 BackgroundTrans Center 0x200 gGuiMove vCaption", CSR Tools
Gui,Main: Add, Tab3,vTabs x-50 y0 w5000 h5000,Main|Functions|EditFunctions|Settings
Gui,Main: Tab, 1
Gui,Main: +LastFound
Gui,Main: Color, 808080
Gui,Main: Margin, 0, 0
Gui,Main: Font, s10
Gui,Main: Add, Button, x7 y+10 w386 r1 +0x4000 vFunctionButton gOpenFunctionTab, Functions
Gui,Main: Add, Button, x7 y+10 w386 r1 +0x4000 vHotkeyButton gOpenHotkeyTab, Text Replacement
Gui,Main: Add, Button, x7 y+10 w386 r1 +0x4000, Search Clipboard
Gui,Main: Add, Button, x7 y+10 w386 r1 +0x4000,
Gui,Main: Add, Button, x7 y+10 w386 r1 +0x4000 vSettingsButton gOpenSettingsTab, Settings

Gui,Main: Tab, 2 ;Functions Tab
Gui,Main: +LastFound
Gui,Main: Color, 808080
Gui,Main: Margin, 0, 0
Gui,Main: Font, s14 Bold
Gui,Main: Add, Button, x10 y32 w60 h28 gFunctionBack, Back
Gui,Main: Font, s10
Gui,Main: Add, Button, x300 y32 h28 w90 vEditFunctionButton gEditFunction,Edit
Gui,Main: Add, Text, x80 y35 cWhite, Functions

Gui,Main: Add, Listview, x20 y+15 w366 r8 Grid vFunctionLV, Function Name|Hotkey
Gui,Main:Default
LV_ModifyCol(1,250)
LV_ModifyCol(2,110)
LV_Add(,"Emulator: License Plate")
LV_Add(,"Emulator: Target LP")
LV_Add(,"Emulator: Short Pick")
LV_Add(,"Emulator: Search Batch Number")
LV_Add(,"Press Apply")
LV_Add(,"Emulator: Skip")

;LoadHotkeys()
;LoadData()


Gui,Main: Tab, 3 ;Text Replacement Tab
Gui,Main: Font, s14 Bold
Gui,Main: Add, Button, x10 y32 w60 h28 gFunctionBack, Back
Gui,Main: Font, s10
Gui,Main: Add, Button, x160 y32 h28 w70  gCreateHotstring,Create
Gui,Main: Add, Button, x240 y32 h28 w70  gEditHotstring,Edit
Gui,Main: Add, Button, x320 y32 h28 w70  gDeleteHotString,Delete


Gui,Main: Add, Listview, x20 y+10 w366 r8 Grid vHotstringLV, Shortcut|Auto Replace
GuiControl, Main:+Report, HotstringLV
Gui,Main:Default
LV_ModifyCol(1,110)
LV_ModifyCol(2,250)


LoadHotkeys()

LoadData()

LoadNSN()

Gui,Main:Font
Gui,Main: Add, StatusBar,, Keith's Automation v%version%
Gui,Main: Font, s14 Bold
Gui,Main: -Caption
SetLoad(100, "Done!")
Sleep, 100

; Fade out
Loop 20
{

    alpha := 255 - (A_Index * 12)
    if (alpha < 0)
        alpha := 0

    WinSet, Transparent, %alpha%
    Sleep 20
}

Gui,Loading: Destroy

WinSet, Transparent, 255
Gui,Main: Show,w%Width% h%Height%
SetTimer,General,On
return



General: ;maintains screen on
MouseGetPos, MouseX, MouseY
sleep,1000
MouseGetPos, TempMouseX, TempMouseY

if ((MouseX=TempMouseX) && (MouseY==TempMouseY)) {
    Counter++
}else {
    Counter:=1
}
if (Counter>58){
    MouseMove, 1, 0, , R
    Counter:=1
}

IfWinExist, Find and Replace ahk_class bosa_sdm_XL9
{
    WinSet, ExStyle, +0x80
    WinSet, ExStyle, -0x40000
}

Return




; =====  AUTOMATION FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================
AutoSkip:
stop := false
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseGetPos, TempMouseX, TempMouseY
CoordMode, Mouse, Window
CoordMode, Pixel, Window

ImageClick("Emulator: Skip", "Skip Button", "Image Search\Skip.PNG", 1, 1000)

; Return mouse to original position
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseClick, Left, %TempMouseX%, %TempMouseY%

CoordMode, Mouse, Window
CoordMode, Pixel, Window
return


; ==================== AutoTargetLP ====================
AutoTargetLP:
stop := false
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseGetPos, TempMouseX, TempMouseY
CoordMode, Mouse, Window
CoordMode, Pixel, Window
if (PickingMode != "DeCA" && PickingMode != "3PL" && PickingMode != "DeCA 2")
	return

ImageClick("Emulator: Target LP", "Target LP", "Image Search\TargetLp.PNG", 1, 1000)

tlp := Data["Emulator: Target LP"]
if (LPStart = 26) {
	LPStart := 1
	LPNumCount++
}

CurrentLetter := Chr(LPStart + 96)
CurrentLP := tlp["Initial Target LP"] . A_DD . A_MM . A_YYYY . LPNumCount . CurrentLetter

Send, %CurrentLP%
Send, {Enter}
; Return mouse to original position
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseClick, Left, %TempMouseX%, %TempMouseY%

CoordMode, Mouse, Window
CoordMode, Pixel, Window
LPStart++
Return

; ==================== AutoApply ====================
AutoApply:
if (ImageClick("Press Apply", "Apply Button", "Image Search\Apply.PNG", 1, 1000))
	Sleep, 50
else if (ErrorLevel = 2)
	MsgBox, something is wrong
Return

; ==================== AutoShortPick ====================
AutoShortPick:
stop := false

ImageClick("Emulator: Short Pick", "Short Pick", "Image Search\Short Pick.PNG", 1, 2000)
ImageClick("Emulator: Short Pick", "Pick Qty",   "Image Search\Pick Qty.PNG", 2, 2000)

Send, 0
Send, {Enter}
Return

; ==================== AutoLP ====================
AutoLP:
stop := false
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseGetPos, TempMouseX, TempMouseY
CoordMode, Mouse, Window
CoordMode, Pixel, Window
If (PickingMode = "SPV" or PickingMode = "3PL")
{
	ImageClick("Emulator: License Plate", "Batch Number", "Image Search\Batch Number.PNG", 2, 1000)
	Send, ^c

	ImageClick("Emulator: License Plate", "License Plate", "Image Search\LP.PNG", 1, 1000)

	Send, %A_Clipboard%
	Send, {Enter}
}
else if (PickingMode = "DeCA 2")
{
	gosub, AutoSearchbatch
	UseWorkLine(A_Clipboard)
	Send, {Enter}
}

; Return mouse to original position
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseClick, Left, %TempMouseX%, %TempMouseY%

CoordMode, Mouse, Window
CoordMode, Pixel, Window
Return


AutoSearchbatch:
stop := false
sb := Data["Emulator: Search Batch Number"]   ; Load section once

CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseGetPos, TempMouseX, TempMouseY

CoordMode, Mouse, Window
CoordMode, Pixel, Window

; ==================== Search & Drag Batch Number ====================
StartTimer := A_TickCount
Loop
{
	Elapsed := A_TickCount - StartTimer
	if (Elapsed > 1000)
		return

	ImageSearch, BatchX, BatchY, sb["Batch Number X1"], sb["Batch Number Y1"], sb["Batch Number X2"], sb["Batch Number Y2"], *10 Image Search\Batch Number.PNG

	If (ErrorLevel = 0)
	{
		X  := BatchX + sb["Batch Number Offset X"]
		Y  := BatchY + sb["Batch Number Offset Y"]
		X2 := X + 250

		MouseClickDrag, Left, %X%, %Y%, %X2%, %Y%
		Send, ^c
		break
	}

	if (stop) {
		MsgBox, stopped
		return
	}
	Sleep, 50
}

; ==================== Click Batch Filter ====================
/*
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen

ImageSearch, ImageX, ImageY, sb["Batch Filter X1"], sb["Batch Filter Y1"], sb["Batch Filter X2"], sb["Batch Filter Y2"], *10 Image Search\Batch Filter.PNG
if (ErrorLevel = 0)
{
	X := ImageX + sb["Batch Filter Offset X"]
	Y := ImageY + sb["Batch Filter Offset Y"]
	MouseClick, Left, %X%, %Y%
	Send, ^a
	Send, %Clipboard%
}

Sleep, 50

; Restore original coord mode
CoordMode, Mouse, Window
CoordMode, Pixel, Window

gosub, AutoApply

; Return mouse to original position
CoordMode, Pixel, Screen
CoordMode, Mouse, Screen
MouseClick, Left, %TempMouseX%, %TempMouseY%

CoordMode, Mouse, Window
CoordMode, Pixel, Window
*/
Return



; Helper: Search for an image and click it with offset
ImageClick(section, baseKey, imagePath, clicks := 1, timeout := 2000)
{
    global Data, stop

    coord := Data[section]
    if (!IsObject(coord)) {
        MsgBox, Section not found: %section%
        return false
    }

    StartTimer := A_TickCount

    Loop
    {
        Elapsed := A_TickCount - StartTimer

        if (Elapsed > timeout)
            return false

        ; Use exact keys from your INI
        x1 := coord[baseKey " X1"]
        y1 := coord[baseKey " Y1"]
        x2 := coord[baseKey " X2"]
        y2 := coord[baseKey " Y2"]
        offX := coord[baseKey " Offset X"]
        offY := coord[baseKey " Offset Y"]

        ImageSearch, ImageX, ImageY, %x1%, %y1%, %x2%, %y2%, *10 %imagePath%

        if (ErrorLevel = 0)
        {
            X := ImageX + offX
            Y := ImageY + offY
            MouseClick, Left, %X%, %Y%, %clicks%
            return true
        }

        if (stop) {
            MsgBox, stopped
            return false
        }
        Sleep, 50
    }
    return false
}


SendHotstring(text) {
    Send, %text%
}
Return


DeleteHotstring:
Gui,Main:Default
Gui,Listview,HotstringLV


Gui, listview,HotstringLV
LV_GetText(HotstringToDelete,LV_GetNext(),1)
LV_Delete(LV_GetNext())
FileDelete,Hotstrings/%HotstringToDelete%.txt



return

CreateHotstring:
CreateCreateHotstringGUI()
return



Contents(a){
;Clipboard:=a
if WinActive("ahk_exe EXCEL.EXE")
{
	xl := ComObjActive("Excel.Application")

	; force Excel out of edit mode (important)
	Send {Esc}
	if !WaitForExcelReady(xl)
		return
	cell := xl.ActiveCell

	oldValue := cell.Value
	newText := "`n" . a   ; text you want to add

	if (oldValue = "")
		newText := a
	else
		newText := "`n" . a

	cell.Value := oldValue . newText

	/*
	Loop 100	; up to 1sec
	{
		try
		{
			if (cell.Value = newValue)
				break
		}
		catch
		{
		}

		Sleep, 10
	}
	*/
	sleep,50

	Cell.Font.Bold := True
	; calculate positions for formatting
	startPos := StrLen(oldValue) + 1
	len := StrLen(newText)

	; format ONLY new text
	charRange := cell.Characters(startPos, len)
	if InStr(a, "SHORT SHELF LIFE ITEM")
		charRange.Font.ColorIndex := 3

}else{ ;if excel is not active
	SendInput % a
}

return
}


WaitForExcelReady(xl, Timeout := 5000)
{
    Start := A_TickCount

    while ((A_TickCount - Start) < Timeout)
    {
        try
        {
            if (xl.Ready)
                return true
        }
        catch
        {
        }

        Sleep, 10
    }
    return false
}




; ===== END OF AUTOMATION FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================


SetLoad(percent, text:="",sleeptime:=25)
{
    GuiControl, Loading:, Bar, % percent

    if (text != "")
        GuiControl, Loading:, Status, % text

    GuiControl, Loading:, Percent, % Format("{:.2f}", percent) "`%"

    ; Process pending GUI paint messages
    Sleep, %sleeptime%
}

SetHotkeys(FunctionName,Hotkeys,Opt){
Switch % FunctionName
{
Case "Emulator: License Plate":
	if (Hotkeys != "ERROR")
	Hotkey,%Hotkeys%,AutoLP,%Opt%
Case "Emulator: Target LP":
	if (Hotkeys != "ERROR")
	Hotkey,~%Hotkeys%,AutoTargetLP,%Opt%
Case "Emulator: Short Pick":
	if (Hotkeys != "ERROR")
	Hotkey,~%Hotkeys%,AutoShortPick,%Opt%
Case "Emulator: Search Batch Number":
	if (Hotkeys != "ERROR")
	Hotkey,~%Hotkeys%,AutoSearchBatch,%Opt%
Case "Press Apply":
	if (Hotkeys != "ERROR")
	Hotkey,~%Hotkeys%,AutoApply,%Opt%
Case "Emulator: Skip":
	if (Hotkeys != "ERROR")
	Hotkey,~%Hotkeys%,AutoSkip,%Opt%
}


}
return



ChangeHotkey:
GuiControl, EditFunction: Enable,SelectedHotkey
GuiControl, EditFunction: Focus,SelectedHotkey
GUiControl, EditFunction: Disable,SetTopLeftButton1
GUiControl, EditFunction: Disable,ChangeAutoClickerHotkey
GUiControl, EditFunction: Disable,SetBottomRightButton1
GUiControl, EditFunction: Disable,Coord1TopLeft
GUiControl, EditFunction: Disable,Coord1BottomRight
GUiControl, EditFunction: Disable,SetTopLeftButton2
GUiControl, EditFunction: Disable,CoordOffset1X
GUiControl, EditFunction: Disable,CoordOffset1Y
GUiControl, EditFunction: Disable,TempEdit1
GUiControl, EditFunction: Disable,CoordOffSetUpDown1X
GUiControl, EditFunction: Disable,CoordOffSetUpDown1Y
GUiControl, EditFunction: Disable,TempEdit2
GUiControl, EditFunction: Disable,Coord2TopLeft
GUiControl, EditFunction: Disable,SetBottomRightButton2
GUiControl, EditFunction: Disable,Coord2BottomRight
GUiControl, EditFunction: Disable,CoordOffset2X
GUiControl, EditFunction: Disable,CoordOffset2Y
GUiControl, EditFunction: Disable,CoordOffSetUpDown2X
GUiControl, EditFunction: Disable,CoordOffSetUpDown2Y
GUiControl, EditFunction: Disable,EditFunctionCancelButton
GUiControl, EditFunction: Disable,EditFunctionSaveButton

Loop {
if IsAnyKeyPressed()
    break
}

loop {
If !(IsANyKeyPressed())
    break
}


GuiControl, EditFunction: Disable,SelectedHotkey
GUiControl, EditFunction: Enable,SetTopLeftButton1
GUiControl, EditFunction: Enable,SetBottomRightButton1
GUiControl, EditFunction: Enable,ChangeAutoClickerHotkey
GUiControl, EditFunction: Enable,SetTopLeftButton2
GUiControl, EditFunction: Enable,Coord1TopLeft
GUiControl, EditFunction: Enable,Coord1BottomRight
GUiControl, EditFunction: Enable,CoordOffset1X
GUiControl, EditFunction: Enable,CoordOffset1Y
GUiControl, EditFunction: Enable,TempEdit1
GUiControl, EditFunction: Enable,CoordOffSetUpDown1X
GUiControl, EditFunction: Enable,CoordOffSetUpDown1Y
GUiControl, EditFunction: Enable,TempEdit2
GUiControl, EditFunction: Enable,Coord2TopLeft
GUiControl, EditFunction: Enable,SetBottomRightButton2
GUiControl, EditFunction: Enable,Coord2BottomRight
GUiControl, EditFunction: Enable,CoordOffset2X
GUiControl, EditFunction: Enable,CoordOffset2Y
GUiControl, EditFunction: Enable,CoordOffSetUpDown2X
GUiControl, EditFunction: Enable,CoordOffSetUpDown2Y
GUiControl, EditFunction: Enable,EditFunctionCancelButton
GUiControl, EditFunction: Enable,EditFunctionSaveButton
return


IsAnyKeyPressed() {
    ; VK codes for mouse buttons and keyboard keys range roughly 1-254
    Loop, 254
    {
        vk := Format("VK{:X}", A_Index)
        ; Check physical state ("P") - whether the key is physically down
        if GetKeyState(vk, "P")
            return true
    }
    return false
}


CreateCreateHotstringGUI(){
	global
	Gui, Main: +Disabled
	GUI,CreateHotStringGUI:Color,808088
	GUI,CreateHotStringGUI:Margin, 0,0
	GUI,CreateHotStringGUI: Font, s10


	Gui, CreateHotStringGUI:Add,Groupbox,x10 y+10 w376 h60, Text to Replace
	Gui, CreateHotStringGUI:Add,Edit,x20 y35 w356 vStringToReplace,%Shortcut%


	Gui, CreateHotStringGUI:Add,Groupbox,x10 y80 w376 h200, Text Replacement
	Gui, CreateHotStringGUI:Add,Edit,x20 y100 w356 r10 vStringReplacement,%Replacement%

	Gui, CreateHotStringGUI: Font, s11
	Gui, CreateHotStringGUI: Add, Button, x30 y+25 w100 gCreateHotstringCancel,Cancel
	Gui, CreateHotStringGUI: Add, Button, x+20 w100,
	Gui, CreateHotStringGUI: Add, Button, x+20 w100 gSaveCreateHotString,Create
	Gui, CreateHotStringGUI: Font, s3
	Gui, CreateHotStringGUI: Add, Text,y+1

	Gui, CreateHotStringGUI: -SysMenu +AlwaysOnTop
	GUI, CreateHotStringGUI:Show, w396, Edit Hotstring
}

SaveCreateHotstring:
Gui,CreateHotStringGUI:Submit,NoHide
Gui,Main:Default


Gui, listview, HotstringLV
LV_Add(,StringToReplace,StringReplacement)

FileDelete,Hotstrings/%StringToReplace%.txt
FileAppend,%StringReplacement%,Hotstrings/%StringToReplace%.txt
FileRead,tempval,Hotstmjyrings/%StringToReplace%.txt
Hotstring("::" . StringToReplace,Func("Contents").Bind(tempval),"On")

Gui,CreateHotStringGUI:Destroy
Gui,Main: -Disabled
Gui,Main: Show
return

CreateHotstringCancel:
Gui,CreateHotStringGUI:Destroy
Gui,Main: -Disabled
Gui,Main: Show
return




CreateEditHotstringGUI(Shortcut,Replacement){
	global
	Gui, Main: +Disabled
	GUI,HotstringGUI:Color,808088
	GUI,HotstringGUI:Margin, 0,0
	GUI,HotstringGUI: Font, s10


	Gui, HotstringGUI:Add,Groupbox,x10 y+10 w376 h60, Text to Replace
	Gui, HotstringGUI:Add,Edit,x20 y35 w356 vStringToReplace,%Shortcut%


	Gui, HotstringGUI:Add,Groupbox,x10 y80 w376 h200, Text Replacement
	Gui, HotstringGUI:Add,Edit,x20 y100 w356 r10 vStringReplacement,%Replacement%

	Gui, HotstringGUI: Font, s11
	Gui, HotstringGUI: Add, Button, x30 y+25 w100 gEditHotstringCancel,Cancel
	Gui, HotstringGUI: Add, Button, x+20 w100,
	Gui, HotstringGUI: Add, Button, x+20 w100 gEditHotstringSave,Save
	Gui, HotstringGUI: Font, s3
	Gui, HotstringGUI: Add, Text,y+1

	Gui, HotstringGUI: -SysMenu +AlwaysOnTop
	GUI, HotstringGUI:Show, w396, Edit Hotstring
	return
}


EditHotstringSave:
Gui,HotstringGUI:Submit,NoHide
Gui,Main:Default


Gui, listview, HotstringLV
LV_GetText(OldHotstring,LV_GetNext(),1)
Hotstring("::" . OldHotstring, "placeholder","Off")
FileDelete,Hotstrings/%OldHotstring%.txt
LV_Modify(LV_GetNext(),,StringToReplace,StringReplacement)

FileDelete,Hotstrings/%StringToReplace%.txt
FileAppend,%StringReplacement%,Hotstrings/%StringToReplace%.txt
FileRead,tempval,Hotstrings/%StringToReplace%.txt
Hotstring("::" . StringToReplace,Func("Contents").Bind(tempval),"On")

Gui,Main: -Disabled
Gui,HotstringGUI:Destroy
Gui,Main: Show
return


EditHotstringCancel:
Gui,HotstringGUI:Destroy
Gui,Main: -Disabled
Gui,Main: Show
return


return



CreateEditFunctionGUI(FuncName,Type1,Name1,Type2,Name2)
{
  global
  TempHotkey :=
  TempX1 :=
  TempY1 :=
  TempX2 :=
  TempY2 :=
  TempOffsetX :=
  TempOffsetY :=
  SelectedFunction := FuncName
  SelectedType1:=Type1
  SelectedName1:=Name1
  SelectedType2:=Type2
  SelectedName2:=Name2


  GUI,EditFunction:Color,808088
  GUI,EditFunction:Margin, 0,0
  GUI,EDitFunction: Font, s10
  GUI,EditFunction: Add, Groupbox, x10 y+5 w386 h52 cWhite, Hotkey
  GUI,EditFunction: Add, Button, x20 y20 h28 vChangeAutoClickerHotkey gChangeHotkey, Change Hotkey
  GUI,EditFunction: Add, Hotkey, vSelectedHotkey x150 y20 h28 w230 +Disabled

  If (Type1 = "Text"){
	Gui, EditFunction: Add, Groupbox, x10 y+10 w386 h100 cWhite, %Name1%
	Gui, EditFunction: Add, Edit, x20 y80 w360 vTempEdit1,
  }Else if (Type1 = "Coordinate") {
	Gui, EditFunction: Add, Groupbox, x10 y+10 w386 h100, %Name1% Coordinates
	Gui, EditFunction: Add, Button, x20 y80 w100 h28 gSetTopLeft1 vSetTopLeftButton1,Set Top Left
	Gui, EditFunction: Add, Edit,vCoord1TopLeft x+5 y80 w120 h28 Disabled Center ,
	Gui, EditFunction: Add, Button, x20 y120 w100 h28 gSetBottomRight1 vSetBottomRightButton1,Set Bottom Right
	Gui, EditFunction: Add, Edit, vCoord1BottomRight x+5 y120 w120 h28 Disabled Center,
	Gui, EditFunction: Add, Text, x260 y80 cWhite, Mouse Click Offset`nto Image
	Gui, EditFunction: Add, Edit, vCoordOffset1X x260 y115 w60 Center gUpdater,0
	Gui, EditFunction: Add, UpDown, vCoordOffSetUpDown1X
	Gui, EditFunction: Add, Edit, vCoordOffset1Y x+10 y115 w60 Center gUpdater,0
	Gui, EditFunction: Add, UpDown, VCoordOffSetUpDown1Y
  }

  If (Type2 = "Text"){
  	Gui, EditFunction: Add, Groupbox, x10 y170 w386 h100 cWhite, %Name2%
	Gui, EditFunction: Add, Edit, x20 y190 w360 vTempEdit2,
  }Else if (Type2 = "Coordinate") {
	Gui, EditFunction: Add, Groupbox, x10 y170 w386 h100 cWhite, %Name2% Coordinates
	Gui, EditFunction: Add, Button, x20 y195 w100 h28 gSetTopLeft2 vSetTopLeftButton2,Set Top Left
	Gui, EditFunction: Add, Edit,vCoord2TopLeft x+5 y195 w120 h28 Disabled Center,
	Gui, EditFunction: Add, Button, x20 y235 w100 h28 gSetBottomRight2 vSetBottomRightButton2,Set Bottom Right
	Gui, EditFunction: Add, Edit,vCoord2BottomRight x+5 y235 w120 h28 Disabled Center,
	Gui, EditFunction: Add, Text, x260 y195 cWhite, Mouse Click Offset`nto Image
	Gui, EditFunction: Add, Edit, vCoordOffset2X x260 y230 w60 Center gUpdater,0
	Gui, EditFunction: Add, UpDown, vCoordOffSetUpDown2X
	Gui, EditFunction: Add, Edit, vCoordOffset2Y x+10 y230 w60 Center gUpdater,0
	Gui, EditFunction: Add, UpDown, vCoordOffSetUpDown2Y
  }

  Gui, EditFunction: Font, s11
  Gui, EditFunction: Add, Button, x30 y+25 w100 gEditFunctionCancel vEditFunctionCancelButton,Cancel
  Gui, EditFunction: Add, Button, x+20 w100,
  Gui, EditFunction: Add, Button, x+20 w100 gEditFunctionSave vEditFunctionSaveButton,Save
  Gui, EditFunction: Font, s3
  Gui, EditFunction: Add, Text,y+1


  ; Get Details
	;IniRead, OutputVar, Filename, Section, Key , Default

  IniRead,TempHotkey,Data.Ini,%FuncName%,Hotkey
  GUIControl,EditFunction:,SelectedHotkey,%TempHotkey%
  If (Type1 = "Text"){
	Iniread,Edit1,Data.Ini,%FuncName%,%Name1%
	GuiControl, EditFunction:,TempEdit1,%Edit1%
  } else if (Type1 = "Coordinate") {
	Iniread,TempHotkey,Data.Ini,%FuncName%,Hotkey
	Iniread,Temp1X1,Data.Ini,%FuncName%,%Name1% X1
	Iniread,Temp1Y1,Data.Ini,%FuncName%,%Name1% Y1
	Iniread,Temp1X2,Data.Ini,%FuncName%,%Name1% X2
	Iniread,Temp1Y2,Data.Ini,%FuncName%,%Name1% Y2
	Iniread,TempOffset1X,Data.Ini,%FuncName%,%Name1% Offset X
	Iniread,TempOffset1Y,Data.Ini,%FuncName%,%Name1% Offset Y
  }

  GuiControl, EditFunction:,Coord1TopLeft,%Temp1X1% `, %Temp1Y1%
  GuiControl, EditFunction:,Coord1BottomRight,%Temp1X2% `, %Temp1Y2%

  GuiControl, EditFunction:,CoordOffset1X,%TempOffSet1X%
  GuiControl, EditFunction:,CoordOffset1Y,%TempOffSet1Y%

  If (Type2 = "Text"){
	Iniread,Edit2,Data.Ini,%FuncName%,%Name2%
	GuiControl, EditFunction:,TempEdit2,%Edit2%
  } else if (Type2 = "Coordinate") {
	Iniread,TempHotkey,Data.Ini,%FuncName%,Hotkey
	Iniread,Temp2X1,Data.Ini,%FuncName%,%Name2% X1
	Iniread,Temp2Y1,Data.Ini,%FuncName%,%Name2% Y1
	Iniread,Temp2X2,Data.Ini,%FuncName%,%Name2% X2
	Iniread,Temp2Y2,Data.Ini,%FuncName%,%Name2% Y2
	Iniread,TempOffset2X,Data.Ini,%FuncName%,%Name2% Offset X
	Iniread,TempOffset2Y,Data.Ini,%FuncName%,%Name2% Offset Y
  }

  GuiControl, EditFunction:,Coord2TopLeft,%Temp2X1% `, %Temp2Y1%
  GuiControl, EditFunction:,Coord2BottomRight,%Temp2X2% `, %Temp2Y2%

  GuiControl, EditFunction:,CoordOffset2X,%TempOffSet2X%
  GuiControl, EditFunction:,CoordOffset2Y,%TempOffSet2Y%


  Gui, EditFunction: -SysMenu -Caption +AlwaysOnTop
  GUI, EditFunction:Show


  return
}

EditHotstring:
Gui, listview,HotstringLV
If !(SelectedRow := LV_GetNext()) {
   MsgBox, 0, %A_ThisLabel%, Select a row in the list-view, please!
   Return
}
LV_GetText(Selectedhotstring, SelectedRow,1)
LV_GetText(Selectedreplacement, SelectedRow,2)
CreateEditHotstringGUI(Selectedhotstring,Selectedreplacement)

return

EditFunction:
Gui, listview,FunctionLV
Gui,Main:Default
If !(SelectedRow := LV_GetNext()) {
   MsgBox, 0, %A_ThisLabel%, Select a row in the list-view, please!
   Return
}
LV_GetText(RowText, SelectedRow)

Gui,Main: +Disabled

If(RowText="Emulator: License Plate") {
	CreateEditFunctionGUI(RowText,"Coordinate","Batch Number","Coordinate","License Plate")
}else if (RowText="Emulator: Target LP"){
	CreateEditFunctionGUI(RowText,"Text","Initial Target LP","Coordinate","Target LP")
}else if (RowText="Emulator: Short Pick"){
	CreateEditFunctionGUI(RowText,"Coordinate","Short Pick","Coordinate","Pick Qty")
}else if (RowText="Emulator: Search Batch Number"){
	CreateEditFunctionGUI(RowText,"Coordinate","Batch Number","Coordinate","Batch Filter")
}else if (RowText="Press Apply"){
	CreateEditFunctionGUI(RowText,"Coordinate","Apply Button","NA","NA")
}else if (RowText="Emulator: Skip"){
	CreateEditFunctionGUI(RowText,"Coordinate","Skip Button","NA","NA")
}
;CreateEditFunctionGUI(FuncName,Type1,Name1,Type2,Name2)


return



SetTopLeft1:
Gui, EditFunction:Hide
If(SelectedName1 = "Batch Filter") {
	CoordMode,Mouse,Screen
}

xpos := "zzz"
ypos := "zzz"
While (xpos = "zzz" && ypos = "zzz") {
	tooltip,Double Click on the Top Left search area
	CheckDoubleClick()
}
Gui, EditFunction:Show
Temp1X1:=xpos
Temp1Y1:=ypos
GuiControl, EditFunction:,Coord1TopLeft,%Temp1X1% `, %Temp1Y1%
CoordMode,Mouse,Window
return


SetBottomRight1:
Gui, EditFunction:Hide
If(SelectedName1 = "Batch Filter") {
	CoordMode,Mouse,Screen
}
xpos := "zzz"
ypos := "zzz"
While (xpos = "zzz" && ypos = "zzz") {
	tooltip,Double Click on the Bottom Right search area
	CheckDoubleClick()
}
Gui, EditFunction:Show
Temp1X2:=xpos
Temp1Y2:=ypos
 GuiControl, EditFunction:,Coord1BottomRight,%Temp1X2% `, %Temp1Y2%
CoordMode,Mouse,Window
return


SetTopLeft2:

If(SelectedName2 = "Batch Filter") {
	CoordMode,Mouse,Screen
}
Gui, EditFunction:Hide
xpos := "zzz"
ypos := "zzz"
While (xpos = "zzz" && ypos = "zzz") {
	tooltip,Double Click on the Top Left search area
	CheckDoubleClick()
}
Gui, EditFunction:Show
Temp2X1:=xpos
Temp2Y1:=ypos
GuiControl, EditFunction:,Coord2TopLeft,%Temp2X1% `, %Temp2Y1%
CoordMode,Mouse,Window
return


SetBottomRight2:
Gui, EditFunction:Hide
If(SelectedName2 = "Batch Filter") {
	CoordMode,Mouse,Screen
}
xpos := "zzz"
ypos := "zzz"
While (xpos = "zzz" && ypos = "zzz") {
	tooltip,Double Click on the Bottom Right search area
	CheckDoubleClick()
}
Gui, EditFunction:Show
Temp2X2:=xpos
Temp2Y2:=ypos
GuiControl, EditFunction:,Coord2BottomRight,%Temp2X2% `, %Temp2Y2%
CoordMode,Mouse,Window
return


return

EditFunctionCancel:
Gui,EditFunction:Destroy
Gui,Main: -Disabled
Gui,Main: Show
return

EditFunctionSave:
tooltip, Saving...
SetHotkeys(SelectedFunction,SelectedHotkey,"Off")
Gui,EditFunction:Submit,NoHide
;IniWrite, Value, Filename, Section, Key
IniWrite,%SelectedHotkey%,Data.Ini,%SelectedFunction%,Hotkey
Gui,Main:Default
Gui,ListView,FunctionLV
LV_Modify(SelectedRow,,,SelectedHotkey)


if (SelectedType1="Text"){
	IniWrite,%TempEdit1%,Data.Ini,%SelectedFunction%,%SelectedName1%
}Else if (SelectedType1="Coordinate"){
	IniWrite,%Temp1X1%,Data.Ini,%SelectedFunction%,%SelectedName1% X1
	IniWrite,%Temp1Y1%,Data.Ini,%SelectedFunction%,%SelectedName1% Y1
	IniWrite,%Temp1X2%,Data.Ini,%SelectedFunction%,%SelectedName1% X2
	IniWrite,%Temp1Y2%,Data.Ini,%SelectedFunction%,%SelectedName1% Y2

	IniWrite,%CoordOffset1X%,Data.Ini,%SelectedFunction%,%SelectedName1% Offset X
	IniWrite,%CoordOffset1Y%,Data.Ini,%SelectedFunction%,%SelectedName1% Offset Y
}


if (SelectedType2="Text"){
	IniWrite,%TempEdit2%,Data.Ini,%SelectedFunction%,%SelectedName2%
}Else if (SelectedType2="Coordinate"){
	IniWrite,%Temp2X1%,Data.Ini,%SelectedFunction%,%SelectedName2% X1
	IniWrite,%Temp2Y1%,Data.Ini,%SelectedFunction%,%SelectedName2% Y1
	IniWrite,%Temp2X2%,Data.Ini,%SelectedFunction%,%SelectedName2% X2
	IniWrite,%Temp2Y2%,Data.Ini,%SelectedFunction%,%SelectedName2% Y2

	IniWrite,%CoordOffset2X%,Data.Ini,%SelectedFunction%,%SelectedName2% Offset X
	IniWrite,%CoordOffset2Y%,Data.Ini,%SelectedFunction%,%SelectedName2% Offset Y
}


LoadHotkeys()
LoadData()
Gui,EditFunction:Destroy
Gui,Main: -Disabled
Gui,Main: Show
Tooltip,
Return



CheckDoubleClick()
{
    static lastClickTime := 0
    static clickCount := 0

   global xpos, ypos
    if (GetKeyState("LButton", "P"))
    {
        currentTime := A_TickCount
        if (currentTime - lastClickTime < 300)
        {
            clickCount++
        }
        else
        {
            clickCount := 1
        }

        lastClickTime := currentTime

        ; Wait for release to avoid multiple detections per click
        while GetKeyState("LButton", "P")
            Sleep, 10

        if (clickCount >= 2)
        {
            ; Double-click detected
            ToolTip  ; clear previous tooltip
            MouseGetPos, xpos, ypos
            Msgbox, Double-click detected!`nX: %xpos%`nY: %ypos%
            ;SetTimer, ClearTooltip, -2000
            clickCount := 0
        }
    }
}

return



updater:
Gui,EditFunction: Submit,NoHide
Gui,Main: Submit, NoHide
return


GuiMove:
   PostMessage, 0xA1, 2
return


OpenFunctionTab:
GUIControl, Choose, Tabs,2
return

FunctionBack:
GUICONtrol, Choose, Tabs,1
return


OpenSettingsTab:
Msgbox, Settings undergoing maintenance (%A_GUIControl%).
return

OpenHotkeyTab:
GUIControl, Choose, Tabs,3
return





UseWorkLine(CurrentBatch){
global BatchLP

if (BatchLP.HasKey(CurrentBatch))
{
    LP := BatchLP[CurrentBatch]
    ClickIntoLP()
    SendInput {Raw}%LP%
}
else
{
    LP := TransformBatchToDeCALP(CurrentBatch)
    ClickIntoLP()
    SendInput {Raw}%LP%
}
return
}


TransformBatchToDeCALP(input)
{
    input := Trim(input)

    pos := InStr(input, "-")
    if (!pos)
        return ""

    return "PO#" SubStr(input, pos + 1)
}


ClickIntoLP()
{
    global LP_LPX1, LP_LPY1, LP_LPX2, LP_LPY2
    global LP_LPOffSetX, LP_LPOffSetY, stop
StartTimer:=A_TickCount
Loop {
	Elapsed:=A_TickCount - StartTimer
	if (Elapsed > 1000) {
		return
	}
	ImageSearch,ImageX,ImageY,%LP_LPX1%, %LP_LPY1%,%LP_LPX2%, %LP_LPY2%,*10 Image Search/LP.PNG
	If !(ErrorLevel) {
		X:=ImageX+LP_LPOffSetX
		Y:=ImageY+LP_LPOffSetY
		MouseClick,Left,%X%,%Y%
		break
	}
	if (stop){
		msgbox, stopped
		return
	}
}
}



; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_background_jpeg(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 85040 << !!A_IsUnicode)
B64 := "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"
B64 .= "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"
B64 .= "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"
B64 .= "cFWyOvFLHfalfQpLcxwKGwBGikke+e1AHPSxTaTdsGY7S3Pv6V19h4b0O/1i4sfEN/DZROB5VxdCTy8dGOUUnKkgkehoAj8LXFtrdhd+C7vTYZhrFuYLcyRqXilJBRkz0OQOvAGa7GX4Mfs8eFrC31jXP219MH2nJih0HwPfagFxglPNDou4ZGQfyqoqL3GYv7IHhzUh8XNSjt9NcvpmgXZ1CGBlUqRIic5PIBz0zXrU9p4I+FvgXU/2kvDth411/Ttc0dNI/ta+8IW1hYXcjMohlDPN5oYsgG5VKkfhWypTgvaNe736Ec0X7q3LGqfGPT/CGuWuk31q8y3dwsKSeZtWIn+Ju5GcAD3ry2X4patbQRatNrkUdxKmZ7SGzVFVeMDecsT+AqozsvIT5r6n1FpV3fmCWaGLyntlDTLMQjBSu7eEbDMuO4zXyDr/AMZ/iHNcQSaV4vvZF2jeZIUbcAT8hYgkjFEq0AUUWPH+u33i/wAW33iTVbhppp7mTbKzchd2AoPoAMViWer/ANoxRyrKvmhtzrt4HtjuD6VyttyuWNim/wCJlhZSzPz+nrSeItQvL/xDLrsyxB7qUNILaERxhgADtReFHsKcdwLmv2z3Ph/7R5WXtJByOuw8EfnWt4atItY83SmkGy7hKE+hx/jVvYDm9GvEW2Upxhu/pVSyWfS7iTTrpMPDMUYE+hx/SswNW4mc3IJ2nAGSOxqKZvOkcL1J61mBP8THmhl0bVY0bbJYlfMA5yj9D+DVe8c2ct74N0a7RflhvHRvbcuf/ZTVKKA0/AfxC+KHg7Tzqnwq1XT7aS+iSPURqFhHcBhGWMe0OCE/1smcdcjPQVgeCtRe00+SPzVKmTIwe/em4gf08QaxDEwVJYX55w5wPyr5/wBc/Zn+Ld5eNqvgz9pG/srluUgvLEywA+hCuDivkvrMrbH6nHLqUV8R9MW9/Ym2M940aoBkSxgOB+VfL8Hhf9v74aKbnSz4Y8aW6rmRLC9a0uJD7RTLsz/wOsZV5SexvDBU27Jn1Mbh4gpt9ThYMAQJkxn86+Z9D/br8KWVzH8Pv2pvA2seBL6XEcMuuWLQQOW6BJxmMnP91jWbqOx0LBW8z6UhQXLGOWXaxyQEXaG/KvC/FPxO8c/A6GPx5YX7eMPAMzBm1G0/eyach7ybMlQOzH5fXFZOtGO5qsDKXwNHsPi7wrBrnhu/0OWD5bq1eJvlyfmBFWPh18QfB3xb8GWvjPwXrsN5aToGWSJwSp9Ditm4S2MJ+1oS5ZI/Fz43eG4rS58RfCDxtHL5Vvcy2lzu+R4ir/JIvcMCFYH6V9ef8Fi/2ZNG07S7b9pjwfpjQXstxHZ+KoYV+S5B4huMDpID8jH+IEZ+6KuknCXNDRmVaftKfLJXR8+fs7eOI/iZ4E/s/wARXsi63oLLp+uIUz5rAYiulxziRQD6A7h1r598BfFq8+CnxHtPiIts0+nyBbTxDZ7f+PiyLZLDvviJ3r7bh3r1o1Ofc8ablSemx7f8Z/hw0qG9imHmxZZWCdhnp6/jXqPinTU8VaKms6FNBdaddQrNbTQ4ZZY3AIdT3BGD+dWmkKUFLVHyp8Q1XX/CDw3MczXFscqzjn5fT2NdJ8VPCV1ok0t2VfDBgQy4Ujk8VNS0o2MWnF6njdo08VkrTKR5ibl56jpVzw7pyawbzwp5gFyoa40pt+csPvxH/eHI9xXnTpON3Y6aU1LQ5XXoVik87aSpPOO1X5bV7ksksTINxVlY8qRwQfpWPUuUGz2r9g7xI+lS6gySkGDUkLbeu1kHP6GsH9kqKLTPF+rabBI4862imBB4G1yD+jVw5hBuKZ6mVyaTifrr+z98UdPttNt7iS7B3KCQfpXzt8NPEl6NKhggu3BCDKg4xXiyvCWh6E4qZ9Y+MvFkOv68NYku1ZUQJCgbhPUgeprxvQvFRmQRTnMnbJ61ftJS1MXBxVj23SPEQljVmlwjccjqK4jw1e382JS6qg67uSfp6Uc8jnlTvudxeLDd6gLuRsIOAx9Kwr3xMsdn9lYKGI+UA9B6VLk2aQXKrI9G0nxVpej2hSWVWHGwdj714/HNdXM4Sa6kZSTkBicUKTjqglCL1Zxf/BRK5v8AXLLw14qs7djY2+ozwTOnKRySIvllvTO1gD6nHevU9e8JeHPH/wAFNe+H95pyzNf2DrZNIB+7uV+aKQE9CrgGtJJ1YmNWNrNHwrbSv/ajMXzhEzx7mnaroXifwn4hufDPjPQbvStUtI0F3p95CY5Izg4OD1U84YZB7GuRxcXZ6E6dzQjClg4GCzHdVW1udygKSTjnNSIkupCLkMP4ScflTLk77kYP3sYHua0pr3kwK3h69u2vGleQg7jkV0nw++Ht5rk/nbG27iCQOOtevN3po1p6M3PD908sgJz1FdfoXwunsnys+4Z6968evfm0PTpyVjqfBSPcW6zTEKoABBNWPDXhu+0+VUZwRuHBrFOVrMqbUja8T6YTozzxMSUHBXuMV2Wh+GYNV09rCeJSrrjOKpQk9jBySPze/beurnw7cXJkVgsqFkJ7+1eq/wDBTD4Aax/wg11qVraM0lhmVSq/fj6mvXy9fvFzHNipyeHfIfk3Fd/a0kcNkNIwX6BiKzoZmsbi5guFMbR3cqupPTDGv0Cm0oI/Mq0oqpJPoWZ4JHjYIeR3qP7dGtz5YcAnBB9Qa6FJM5XHS50Xw91a9s2aN1yrEZ3Va03SUt7aKZSf3zBUEYyXY9FUDksewHJraCe4QV+pZ174yeMvAPjsal4K1aSSJILdr/S53LwzDaQ3H8J4HzLg9OtereBf2MvEWnawnxc+MekXtjYvYr9k8KxW5Fze9cNcN/ywQ5ztHznAyV5FeTja+Xv3arT/ABPYwWHzfnUsOpR89ke8fs1/tx6vafC228Q6JI93pKhxPp0ozLbTLkPF+B/AjkV4d4ovtXtLkzaVZpptnAVEOmWkQjjhVckBVHpz9a+Zq4LA1qt4bdj7KOMx1OhatZy7o+gF/as1/wAc+NIvH+j6pJZ3tqR9l2NnyR3XHQqe46HPNeA+Gtc0FbxJ4r46ZPIMvIYy8Eme+Byp/SvQpU6NKHLGOh4VepXrTcpS1Ptrwz8Rfh/46shPqqt4f1F3LTmyt/MsJWP8QjB3wHOeF3LzXgXhXX/GB01ZvD9vp935gSK3K3QL3EjsFRQnByWIAAyTmpm6MNb8tiI+1a1jc+6Pgj4Z0fwx8Otf+LU3xc025Tw3qiW+maHYaynlR3c9rIqySbtskLhpo2URruwjnpXjmta1pUd/Z/Dvw3cWH9l6dcvLLc3unizv/tTxpG0UyLl5ihjYb2B5kKr8qgV53NDFVOaXwry3FyVYWauvI6PWtFl1hm1m08V6XJcxDf8AYbO4k/fP1wu9cLnnjP41o+CD4DsGeLVLq4v72IkXFqn+ixwuDgq5YF+O+QO9bvDYSS91WOuljcZT+NnhX7WP7fh/Zs8I3NpoaR3/AIvaydrTTWOBFtUZdx2Vdyk9z2r0f9p79lz4TftI+F7a1+Imi2OnarpzzLouq6Qivf2ocHcDJk+bEQQNkmVOBwMCunB4PAKd6ybRnjs0x86dqLS/P5H53+CfizefErx1f/FbxP4ruNQ1DX54WmuLqRmIkWBVZBknYoJOAOBWf8Wv2JP2hf2WdbgjvNPbWvCE+pqbLxPpEDGONWIBFzHy0Bx/Ecpx97tX0sa+HhTtTskuh8zCFepW9+7bP0M/4J/eFG8VWEfjrWGLtdqsdrIyjd9nQkRqM/w4JIH+0a6T9inUNP03wTp8FmyCJIYxGsZ4AAwAK/Ps0nzV5SfVn6Vl9N08NFPsfafgzw5YW9qp2ZPAFQ+BPElvPbJmQAHsx6V4UuXmudrvbQ7CWwjS3G1Ce4ANPFykkIaKTI25OD1FJtWITlcwNae6cnjgr36VavZYplKbCTk8VhK0upofMn7adhdp8JvEV1GW+XR7np1/1TV6D+0d4Pj8R/DrWdL8snz9OmjK/wC8jCrwjUMZCT7l4hc+Gml1R/NneJDNI8sTA+VIV+XGCAcZ4qraWtzot/Nol9w0FxJBKp7MjlT+or9si04po/FJu02n3ZpW90ZPD08D5LwTK6YPRWBU/wBKpmZ7Uy2yEYddjZHXnIpmbdzZ0CyuJUFnZLtZlDXMzDARfTPr1q74ccSQJFDECh/1rOx5prcRpxPbafD/AGfou13I+ebbkge1aGn26WqlYTEm4/eSPn9a1ewDdPgkgwCrSM3VgO/vVtbVmGZLtmGRx0zSW4F21klVghhC8Uy0SNGIKDpwdxJqwNmwaMAMkAXPUnqTUdvuijHn/wAQyoIqouwG5pUxV/Mkcg/wiqthKpwN44AABNacyFzI9a+DXiBrDxNav5uW85QM+5xXLeA9TmsNRhvok3mORXRFH3sHOKqMkLd6HzN4v02LQvidrWiIfls/El7bJ/ux3Uij+VfpN46/4N3fE/ibxPe+NPD37aPh03GsX8uojTb/AMG3AaGWdjK0ReO4YEBmIDYHSvNlUhztXLUJb2PzPshYG4ke9YoBMwVmU4HNW/EtjH4R8Yax4T1KdDPpeq3VlPJEcq7wzPEWGRnBKEjNP4kIuWKR3l/b29s2Q8yKuR1BIpNC1fQLC8hv3uhK0DhvK2YyRSs1sC1HeI4PFFzr07W2l6iunNdMsV0LWTyCejBXI2k/Q16vefHHwRon7Ovh/wCHkHiOG71Vpbi51WzjG82zM5KIePQ9qFKTdiuU8jv9NksPBV3D9qlcPeRxQW05JRCRuc7M4GcDOMZqx421JLvw1pl6iFY7u9llADDcdqhQSOo/GqJNz4kfHzW/ihpsfhy80s6fodvaW9vpWhpqUk8OnrEoAEe4DqRu5GRnGcAVwEdxA8gjWQHvyP1rWderOHLKTZEadODukakxgnZ550V2x97GT0xUEdxDGmGbLY4wKzbuy9yheajexSpc2kzIYuFx/wDXFLPbzTkrHGdu4fhmkBo+GA9xfPdXLORsBbavOM8kY64yeOK2/AWh6l4f1S01vW7B5NOknUCWPo7KdzR5Ixu25oAq61BprXNzFptzI9uJT9lkmQK7rjqwHQ0a9Pp+naxelyVQ3TPDHIQWVDyAcexHSgDW+HWow2t9FJNDv52qQeQfasLQNVMF00lucKDuTK8ChgavxJ0rSYvFt49lIIHkgSeSFgcu7HBI/LNUfiJ4gl8TXNvOkaibTrcQSSJwZCTnJPfHAHpU8oDGhlV9hZt5PT2xWPp/iOe0kWKW2SVd2N+SHX15o5QO/Dxy+Ap7DV28s28kMkZdTjZlgT+tS+EL7SfHfhvWtKiuyZxo8wMc3LhQuQf++scijmS3Gk3seWyXuoWg/wBAJ+ZjvC9BjpU+q/D/AOLEqxJo/gDxIrqCZlj0W54BAKk/Jznkik3Fhys/p80jU8KoLKefWsbSb1UcfJnHcV8UmfsUYs7qw1ZlA8oDOOhrL026yw2DORk1LaZaj3NXXbSw8ZaHNoHifwRZ67YyqRNp93bLNGwPX5XGKuabOIyDESjeoPNTr0KsfOXiL4SaB+z1rw8TfAjxpqnw9hlk/wBJ8Pa9aSz6FcDuqkhvs4OT8vKY/hFfU6TxalaPpmr20VzbyoUlhuYxIjqeoKtkEfWlySfU0jiHDSWq/rufJvhC18Q/CTxwvxO+DHhiPTJNWlM3ijwJp90JNL1pcEve6W4OyObqz2+QH5YAPnd6b4q/Zg1z4Z63P48/ZrhiudOuZvN1n4a6pc4srs8ky2Ejf8edwDghciJsY+QncFCFSJt7ajKNkeia3o3w1/aY+EF1oWtWKaloPiGwaG6gdSroCMEHoUkU9DwQw7Vx3wx+KvhSbVLjX/DtneadH9qFv4u8P6jGY7mwuOgndDyDyA/qCrgkcm41JRZyVKEJK6R+Rn7a/wCyL4v/AGWfivffDTXnlubG733PhzVniwL+2zwT28xeFcDvg8BgK/ZX9pL9m/4ZftMfDe48AfEnSBcWUpMmn6jbgfatNnwQs8D9VYZ5HRhkEEGumFeSdzzqmFU07H5D/wDBOvxX4luLvWvhZq3jSCTTLGxSXRfDt9xcQSGRvMFu3V4cfNsP3O3BAEv7UX/BP/44/sreJ5PFGoafNc6RaagE03xZo+4I4b/VudvzQORwQeMjGcEV1RxOhyU8KoVLO56R8Yvhc+u6fc3Fhb8ICZMZJA6cf/qrg/Cn7Q/x8m8Om0j1/T9UnWH5Tqumq7EjoC0ZVqf1uCep0Sy6c43R8+eJdK1X4f8AjSLUpYebO5EkJVDiRR1B+oJH410fxH+KHif4kTy2/jH4aaVY38EhV5bGeZct/utnHr1rR1qUoWbOB4GtTnzIr/GPwdBp+q2uvaNJmy1e3E0e1cfOMZH4gg/ga3/DfiXwNqXwos/DvifS9XEsGpGGS4s4vtLWjkk72iA3GL3XJAOcYBrz9U3Y9SVH92pNpFP9nFf7K8es0x4n0+RBz1IIP9K2rbwbdeCNSi1Bp42KkGGePgOpHH0yDXNXvVhZ7m2FpyoVOZrc+nPhd4pkDrGsmcHDYPSs34E+Hr++0+G8nRn85gxA6tg549u1eHVSO+covY9+8JWgvoknWVt7DIbbkAdya6HwRolytlBBIVjYr+9GPuk9ErlvJOxk52R0vh6K6it+JFC/3gvQe5/wrqtM06AaWlteXQeVBlUkUAD8utaRvbUwb5mc1NB9pJlEXGTtNS6wrokhaUuGb5FUd+/QVQipIfLiQyThSW421mXk9zBDJd3LKsa/6oOMFjRZvYe5uReNToMPky3X7s9QV5xXIeBNA1b41/EBPBOh2bvbQqJNYvVBEdnBnu3Te3IVevfoK66eHqT2E7H0b4v+C3ws/ao+EWkj4k6IrXr6Yjabrduuy8smK/KyPjJHAJRsq3cc13Wlrb6fYw6bZxqkFvEscKKOFVQAB+Qr2vZU50lCornO4XZ+WHxC8CeKfhH4+1T4beMrcJf6Tc+W0iKQlxGeY5k9UdcEehyOoNfV3/BUP4V6dqvgbSfjdplrtvdDuo7HU5oxzJZzNhN3qElK4z90SN614eNwbo+9DYD4+sHkutdt7dehYE8dhzUnhdkXV/tkmCqIcEjucAf1rnw8G6l3sJpvY92+GNpZadpCtIAGJzt9c1leE9S22ahJBnAHHrXbWktkddGCa1PQra6h3q8L4IPzVkabdMhXzJRz96vMqO7O2FO6O90KaGW5RflY7vmDDgiqfhS6hkkFwwGMcHNcsqlmOVJ2PXvC2n2EUYaNVwFBrl9I8XwWabC+NwwOeldNKtE55UZGd+1d4PsfE3gG4aS2jkCxFZNw6qRgiqXxV+IekXOhXOiR3gmmmjKrFGcnJ9ewrtp1UtUZ+ylFWP56v2v/AAbJ8L/2hvEngm3+W3N6JoMcDynHX9DX6oT/APBO79nrx38Z5vjh8YvDqeJtYMcaWulXpzYQBd2CYh/rm+b+PI4GBX0+GzzCYbDctS7fkfJ47h7F4jGOVKyi+5+bv7NP7GPx8/a71u2Pwo8KiDQ43CX3izVt0OnWwBwdr4zcOOfkjzkjBZa/bXwl4ZmjtrfSdP06K1soUCQxQRCOONRwAFUAAY9BXFiOKKm1GFvNnbQ4Tpx1r1L+S/q54v8Asn/8E+Pgx+zZZWt/DC/ifxLCu2fxLrESExMeotogNsCjpnlyMZY19Ow6bovh+MMZY8MMOT1z714WJzTMMV8dR/kvuPfwuWZdg/4dNeu/5nLXXwn8N6zBJFfWccm4HIlXcD75NW/FHje1i329sD8vTYfvCvOfNe7kdy5nsfI/7a37KGiaFpMvizwnpqRPEu6dFAAZfYCvaPjVLL4k8FXsd1G4R4WG1x7GuzDYmrCVkzKpRjUi7o/LPXLCPSLtkJxCzYXn/VsT/KtL4qJY6TLfTalOkUMMjb2boACa+qwtac4qx8vjaNOm227Hq/7LaaIlm/jHxxp2kziG8Gm+HNP1vzEXULzY7Ga2ZcbpoWEWDu+UsW524rwq3/ap0f4o3mneCDaXGh2ehW8Y0O3ttio84jCyzskeFMh25D/fKn5iTmrr4PEVfekuVff955tPG4ZO0ZXPrn4NapLqekajfa5BcSeI/BmuR61fXMi7pbq2LqLlGPUspVmAPXC461wnwM+KHi3WfE48bQ6rb22qSWgtrqYW6+XeQ4CsJlb5WyMZzjkCubEUHZcujWj7Hfh6sY67rp39D0zxZ4q17wz8W9e0WS7bZNqLXsUuOLiG4AlRs9x8xA/3a5fX725sNZs9R1fXrfV4xpqw6dJZkKLdYi0X2Vi3Qxse/VSPrXbQv7LlucOJalW5ktD3nQPjnY6ppMXh24ks1S0IN7qLwgMjdPKV1HKjv1yfTFeE6HHrmkNLbaXrMaHY486LBV+Cc7TnBPP41rGElK5yy966kfWVtHoU+kwyabf2V+12sir9hxL50W0E7g3APJG0jJGfTn5y0f4t694Es7OTTdUaGM2gcPHGrfOeGYbgcn5cZ+tOVNPUOdpKx7To/hOHwVo2peM/DlrFbWdlNCTaIAiu8j4KKo6EDJwOMCs68+KMeuwSSXtrbLF5sUk62J/0eG7miDkAHOcr2HTkcVxVcBRrP3jto5piMOu57d8Lfidb6hDEyOUyudrf59a4H4PeHW1bxDY6Ojy/Z7xGIdRgK2wsACehz1HfNfN4/LFh7zpu6PpsBmUcSkpqzPqPQ/Faz2wEbAjbxzXjOra5r/wv1JXv5JLnT94R5jkGEk4BPt714PvXPaWHjNXTPaJtVjL/ADL9K4/RfFVvq9mtzDMSu3r60yXSsdF4osrfVtGkidtx24O0ZODx/WqJ1VZNJeRCQQPmBpxezXQHFs/m1/ac8IT+Af2lviF4PlhZP7N8a6lHGMf8szcu6f8AjrLXpv8AwVItLS2/b++Jq2oCpPq1tcBQOhaygz9fmBP41+xZVWlXwFOb6o/Hc3oRw+Y1Ka7/AJ6nz7LKzspHWmuPnwK9Jnms63wwzQ2wMhQAnjLjNHhbR5hAt20aRIDlpZyFP4A04iOkhucAYQEYHJHeo4NU023kIsjLdyd2x8q1QF+3S4lIcsFB/iI4FVYvt2pKftFyVXJPljigDTi1CGCVoLCASyAcux6f41TWKKABQoVvVTTuwNiBmZRJPKS5HQngfSqlk7uFYZP1qwN3TZlidSEye5qLTWZjwvJxn2q0xNHovwd0yfxB8QNG8OwAlr/WLW2TH+3Min9CaZ8FNYTw78T9D8TMfl0u9S9c5wAImDH+VNxbixJNWP2R+Imm+MdU8Fa7pHw91tLHxBNpVzFoV9KPkgu/LPkuwHVd2Mj3rTt7hJ9Ringkyksyuhz2bkfpXgbVHc79eQ/nI8WR3l74i1WDxpoka66mrXi6tIsnH2zz5PO/Dzd+OTxVr4+adq0Px/8AiBBaW8oSPx5rKqyZ7X0wr0IXscT3OQu9LurBg0sCgZ+VlbrVoQaxcFfttvcMqggExHGceuKsDNfLPweM5+WpHguoF3SWUy9xugYD88UBdk0MztbmGWdiqf6tSc49cVHBibjKRnvvbAoAnRbZULEfPkbWBqHynV8M6ke1AF6Of7SxVRtKjpVQSNE3mBsYHNAGppSwSW18b27COsIa2/2nVgdv41T+aWPdjII7daAOj0b4htHoU/he5do4pJ0njXzCVWRT1Huc4965OZQjEBRz6igDrtQ0+01y4GokBQ4A3buM1ieEvE2q6HdSxWVp9tjmhaNrV7bzsZHJRcHDcdR05oA1bfQ3tJDKk+9UOTzx+dZ03ifUL9Vt2gEMS42RID1989aT0QFhbJ4JrqacBhOc9ec0yKe5upmVFd8dcIT/ACqOYdmZd1YeXNuBIyeoFdl4Y8GNqd19o1bfHCP4W4LH2z0pc67i3Kfw7tRpfjjQ9QSe4gtZ7o2+qT20e9xauCJsKeGITJC9yor6A/Zw/Z08ZftIfEC1+EvwotNEh1H7JNdpLr2qrY2VtEm1HmmnKttAMijAVic8Dg1y18RTpR5pySXmdNPD1KkkoK7P0w+EOrQ6B8K/D2jaJ4o1PULODSYRa3+osq3E8ZQFWkCAANgjgele1/BH/gnrpNr8MtH0fxp+0zoct/YadDbz/wDCLyQTQKyoFP7yV8t04O1a855pgF9v7j0o5ZipK7iaGkXSjBZvwrL0i9VmRi33jjNeJdo/TTtLG5uVw9rMA3bdVXS5GYAJgH60gN+DxD49gXbBoFneqOcCXa1O0aaaJkUg8jJ9DRsBLF8Vb7RkB8SfDzVbdR1ns4/OUD6Dmtux1CSIbkOPXaannkS4pk3g/wCLvw88WTnT9F8SwNOOGtZ/3cgPptbBrO8UfDn4d/Ey38vxb4bikmH+qvbcmG4iPqsiYYH86vmDlp9dCn8ZvhhZatqMPxG0dEtdatIzAt6QRFfW54NndAfeQgna55QnI4JB5ySX4xfs/NJFrsd18QvAE0ZS9aK236tpUeOWeNf+PqMDq0Y3j+6eTQ7S3RpC6fuvQ7v4O+JX1HTZvDGoGTzrDaqrOP3hhYZTd7ryhPqhryXwR8QNO8OfHjSLLTtcjvtM1rSpP7M1JZty31puEkEgx1Zd0kTd8ovrURqRjLlZVfDupDnie2+INB0bV7a48NeJdLhvtPvF2vb3MYZDz0INaOu4lshcLgspDDFatyRx6SjZ6nyp8cf2HfhP4e1Wbxp4H8NQ2O9S9xa26/IWHcen0r6V1fS4PEOjyWU4DGRSAT6UckZpmtOrOi/I/Hf9sH4V2+keJovGml2aoJ08m8VBgB1+4/1xkV9E/ts/C0WUep6ROhVGUtEygdRyKilLldmdFXlqw54n5/6vJeWVhNd2UssbKMu0DlXQjkOMEHK/qM1qa1Ztpd8Q42owLDcOPpXSnbU4p3kuUZ8FPFfxP+LXjbTvBOoeI5Ly3tZRNeO0CM5hXou7HAJK8nnGa+gf2Jfg5Y6EJfEYsR5mpTLL6GOPHyL9Op/GuTH4+jThywWpnhsNiHL35aH0/wDBnwHHGbOCOxK+TbZJHAr0P4YWdvLCZYWIAxuYHBz0wB6V85KbkzskuU6DQPDzGUG5kVN2MbYxn866CexspLAykuflGXxSVjFtvcpTTXOnXCySSM5UhMqnT61FqY1CxkVLBVffguWHGM9efar5mTaxHro8m48y5YEMpbYoB3n61FryXV+tvY2MPmSTTrFaogI3ux2jn0/kBmtoU3U2B7HCeKr3Uda1KDwt4fQyX97IVhTJwi93b0UDkn0r6P8AAPwB8KeAdEu5Yw15rWpoBqGqSqNxHaKMfwRDsvUnk5NevRwnLG7FGabsWfhf4E8MfCjwjZeHfCwEquBJe30ijzLudh88rkdST0HYAAcCqXhi8msbqbwrqDnfDJvtyx5YdxXZBxgrGk42Wh0/9tRwXiwSEgSlgh/2hXM+M7xrS80tFbb/AMTULx3BXmhtozsjoPGfhzw78QPB1/4J8V2gudN1a0e2vYScZjcYJB7EZBB7EA1We9VIwhk4KqF5602k1ZmLSSZ+Y/xD8Nav8JPG+rfDHWnaW60jWHtpZ2Xb50aHKS/8DRkfH+1XZf8ABR7/AIlX7TWqTDcPtmlWNwpI6kx7Cff7mM+1eXKhCjUaRpQvMs+AddSSBNswwF6CvPfAPidraxRmkzhcE57+lc9ZJM9OnTdtD3AeI7eGaOOWQDLc+1eZSeIBewsZLsx5++yn5h9M1yeyUup025Fdo9ps/iZpNsXsNJdZ5YwPN2twhIzgnt9K8K0zxLaadG2kaKnlQq5diDzIxOSzHuazdCindke/LZWR6/q3xDvb24IvNVdY+gSA7fwryhNe2MLiS4Y7edvvRzwirRQ/YtrVnrWg6pNqTiGxRF3MOTyT9a840T4lXGlyYtxwWLAHuTUuTY1TaPc7GPSdDjF1fSLvC52v6147ba34m8YX6SpcMoUAOuTkrWd2JxT6nrOrfGmzt1FrpUQdyueB36CuW0XwdZ6di9vpF3Ln5T945qW9RckUa9t4n1zxCw/tAsBuz8jHnjpTNPv7C23tAR8pyqMep9aly10B2Rs6VpCkgKh5OSW7Vh6/8RbHw5p8l/qE6xCMBiegx3zmrgnJ2S1Jk7LXQT486vonhf4b6hqd8Y4FihZi5PQYNfm1/wAFF/2+Ln4pandfB3wBq6jToJfL1a5t3J8wjrACPwyfw+nvZfkeLxLU5qyPnsy4iweDi4wfNLyPnf8AaF+LMnxI8VXuneHpydKhunIlU8XL5+97r6e4zXDSJBaWbzDbkjHA6e1fdYbB0cNFJH57jMyxGOnzTfyRz9xdXNpfRS207RSRyApIjYKnPWoLgyLdea8Zwrgn8662uZWZxJ2d0fUvwN+MNkNOSDWNMhuUgO68hmJ8uUY7gE8ZwQccGvB/CHiybTdTi1CwnKMvDLnG4ehGeRXnYnL4VFdPU9XC5pVo+69UfcWqfEjwH4j0lLrRb53v5ECx21xcK80Y4AAkU5cAADEgY+9eJ6Zd/LF4gt7CNIngz5SZCk7ccEe/NcMaEqOl38z0p4p1rN7HuiN4v8EeJ7a01KyMlzNbRSSKjZEazA+WrFsBGJXqTjnGa5zSfiPqHj9BJ4jvwbv7DFa3MkmAlzEjZQSY6nPcioTqpe8VUVJtcjPXfC/ie01nSLzSL3TYZ74uq26ygEJIHy21VJ39xjkda8qvNcvvCGsxP4c1LY3lfPdRMVUseqp0KqBkZGM9qp++rmbXJK259KxXOmww6xrMFtHPpun3CmzWdsJc3ksMceGU/wB1BK2BjBI6V5T4X+KlhqnhseHyqwxGaKYRgkjz1RkZ/wAQcfhWUqSbRtGcVfQ9l0L4q6zpPhqA27i3vo9Qa4tpYmOYQgxj1OSFI+nvXLfDDw3e/EXxhY+H9KKsksgN1KchUhBG9j6cAgepNcuKhRowlKf4nVhfa16ijTWp91ap9j8W+Gopb6BG+1WyPOjLwSygkfTmqmiOI7FYFT5FXC+w7V8DNpzbPu6NNwSuzL8F6XJ4XR9HSR3gVsw7+SvtWlPKVnxnk9Ki6NW2xur67Jp1tNiTOVOAfpXNfEPURYwISxKseT0HqfpwDWlOKdl3KaVr9j8Pf+CimtJrv7cHxJ1GG4LouuRwoc8DZbQqR+BDV55+0B4pbxl+0D458Vq+5b/xdqEsWDkFPPdV/wDHVFfsOWUPq+ApwfRH4rmuIWJzGrUW1/8AgHLs7eaHBwRjBpitkhpAcDriu887c6bw/aC+YPfzPKeNoZyQPwNaHhuy025gX7BqCSNtHyScEGqiBtWy+TH5UaIB7JinxaZrkUO97MMo/iTk/pVASxpI7lVYlScgY70WE1zFLtMR3EYwR+tAE4g8wkt8uO5q+vMWGjHI6E0AJDbC3QKQrE9DnpTY1KYZTkdl9KrmYGjpDyQynKj8qZaXBVwxPJ7e1UmB1Gh3cNppOp302FX7DJHuPI+YYNZWq30th4caBGwZ4d2P7wLAD+Rq1OSYWP2B/Zc+K2r/ABF/ZR8M/E3TbH7fq1v4ZkRLPfj7TeWivGUJ9WaIf99V4J/wRm+KE2tfCzxX8KtQuEMnhzxNHfWChjkWt5GrHP0nSX2wRXlYqKjUvY6acnKNj8oPiR4p8Q+MPiT4k8Y6lpVzpLaz4k1C/lspYmBtZZbmR5ITgcFGLKQccivof9of4QXlx+0b8Q7lpyEk8dam8aW6DIUzsen4k/jTVZpLQh0pNnzHpHiS/wBJvGaC+MrSRMg8xjjkdcGvZ7/4K6NIoGqaXNdbWyBOu3n8MU/b+QvZT7HjFv4v1tCEa+d0HVJfmBr1G6+DXhpj5NloJtmHeNmPPrg0/bIXs5Hlt5qt9dwSXR0iKSJADKyrlUBOMk9ua6/XvBN54ThdZrGaeCbhwkR+bByN2O2eaftUTyy7Hnj3kTMVW0AOf4Wr2P4TfDHSfiHY/abzTrS3O8qyzgoWwevOMUe2iNU5s8y03wg+rRCa41AW45wNhbjOK+l9I/Y48Kaq8UNr4ht7GSR9saDW1UZPcKT1o9o3sN02meAW2l6JpCC0tC8zkbW8wjJ9xmrfxC8BW3h7x1rml6VI12ul6nLZSy3Egy0iAc5TjB6j2q1JtEOLRkeKdG05bUX0NykcoBXylOd2O9UEg06wnIu5LaN88LJcDJ/CjmFsbPwJultfizoc5jVsahFnI45OP61Y0ObUILmK80Syb7REytBJbxZII5BBAqZO6sBHf24g+ItzDLsVV1KWNkA4A8xzVu71nQ7PXf7S8S+FzPeGcPeWsk7wlz3Y4H3jnPvU+9aw+p1+n21hsEduE3j0I6VhaVrHhnVNSP8AYkDmRm+W23EOPbA61Ps5FXR1ttYS3FwsW3OT0rPg1G70x/NOm6jCQflG0nB/EVnyscZW6Hs37OHjPxp8D/iLB438EwWBu7i3OmzR6lbtLE0E0se7hWUg5RSCD2PBzXnHhT4u2sXifTtJk19VuH1K1JgngVnIM6A8Agge+K5cThKGLpuFWN0dVHETpT5oOzP0R8aePfHOiX4tLrwZ4cvHVnUzNJNb7gCMHGHz+favYvEPgzTdfvC1zYwy+W52kRg4yB/hXgvh3Kr/AMNfe/8AM9WGa4y3x/kdpo18yxoikZHcc1R0KWX5NwH5VnLY/RTt9NvLx4BJbyKGz/GuRUWiPuCo0Z5OcioA6PT9f8fWh/cWum3CjoJQVJp2mOCFRmIx1yKHsBoweMvGaENcfDSyuOOWttR2n9auacR0GTkdfSswJbT4kwWjA638O9esgPvSQRrOi/8AfJzWpp8skWPLkIPvVcqC0expaB8RPBOqEJZ+JIVkzxDdAxOD6YPemLYabfnbqGlW83+1LECa1joiJQV7o4r4zfAHStVtP+Ey8E2KWup2Ny2oWItQAhuCCXAA4HmjhscMwVjyMn0fSdF07TGI0wtCrdY1YlT+BpOnGYRrzotyT+RifDTxvZfED4eW+t2suTJbAsD1U4HB9DXGeCr+T4d/HrxF8Nr2F4bDVR/aejvIMI6yjMir/uyhxjrgj1qW2o8rKlTjJKrD5nodtOu1Yw2CvoazJbz7LqrwuR8uM81F0huHMtD53/4KK6Bbx+Fm8QxINyx5cDuMVV/4KAeNNNk8MS6LcTpte1ZWUn1Faq0/UiF6Wh+eqeDbLxdqkFpNEWQ3Qc4/u5yR+Ndl8BtBur+f7Y6ZLXBWHnOFBIz+Nc+KxSpRsmdFKjGb5pHu3wd0FdOso4ki2AKANg/Kux+H2jy/6Owjwo/1mOce9fPTk6juzaVlsj1D4bWs8U6tBEE8rAkBH3vetrwrpUtwyv5ISARYPz4+bPUmmoqxyyd2bl1cvJI9uYdqtjMpPyj2x3ql4hhksrZra3kBMrZJdyQg9fpTUbsmyFnuxZxNE0qSFlyX2jcMduelczfXbSI7uWcl8M7fxDHNaKDTSRM4pM7X4M6Tqeua7L4x1Pb9ntroQ6fCeSvGWb8iB+JrtvhjpEWkeANLjQq3mr5rbB3bkDPsMV7eHpRhC9idkzu4pGmTAbqOKWxhLDPQCu7mujjvdnFfE3Qru1u4fFGnp+8iIzjvg9K7TW0sBpbxajLGqE/ekYADj3qXE2hUktDyb4ka/a3F9oEqOV8+Yzkeyrz+RrE+LXij4b3ur6d4X0HxLZNrFrbzTtpsd0rSrbMyK0u0dFDEDPqacoVGuboCnDncLnQSaxJcPBGHIHysfyrltP1r7TdtcDhMAKAaYpwdj5W/4KmXg/4W7oF+y8HwyqyHHX9/JjP0/rVD/gpZeWeufEW30WB83Vt4MWY89MzysmPqUIrzMZJ0q8W+oYaahUaPEfCuuPDP9jMgwcYGelYnwlni8R6qkakkyIGGa5K7SVz3KTTR6K1600BjDEkcmr3/AAjE8V01t5ZQrgHntXnurfQ6dLGbpYuBcDLHk85ruPDvgGSeDzJ0JbORzUupcXumLBZTXRMSnB68eldkvhCTTds3k5ZTx70rh7qWhj6T4aEEy3F3gqFyFbnmtHU5LkSpHFGe"
B64 .= "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"
B64 .= "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"
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}








; ===== LOAD DATA FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================


LoadData()
{
    global Data,loadval
    Data := {}   ; Clear old data

    ; Read the entire file as text
    FileRead, IniContent, Data.Ini
	RegExReplace(IniContent, "(\R)",,count)
	;MsgBox % count
	a:=40/count
    if (ErrorLevel) {
        MsgBox, 16, Error, Could not read Data.Ini!`nMake sure it is in the same folder as the script.
        return
    }

    section := ""

    Loop, Parse, IniContent, `n, `r
    {
        line := Trim(A_LoopField)
		loadval := Round(loadval + a, 2)
		SetLoad(loadval, "Importing Data.Ini (" . A_Index . ")...")
		Sleep, 0
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue

        if (SubStr(line, 1, 1) = "[")  ; Section header
        {
            section := Trim(SubStr(line, 2, StrLen(line)-2))
            Data[section] := {}
            continue
        }

        ; Key=Value line
        if (InStr(line, "="))
        {
            keyPos := InStr(line, "=")
            key   := Trim(SubStr(line, 1, keyPos-1))
            value := Trim(SubStr(line, keyPos+1))

            if (section != "" && key != "")
                Data[section][key] := value
        }
    }

    ; ==================== Hotstrings ====================
    Gui, ListView, HotstringLV
    Loop, Files, Hotstrings/*.txt
    {
        FileNameNoExt := RegExReplace(A_LoopFileName, "\.[^\.]+$")
        FileRead, Contents, %A_LoopFileFullPath%
        LV_Add(, FileNameNoExt, Contents)
        Hotstring("::" . FileNameNoExt, Func("SendHotstring").Bind(Contents), "On")
    }
}


LoadHotkeys(){
	Gui,Main:Default
	IniRead,TempHotkey,Data.Ini,%FuncName%,Hotkey
	Gui,Listview,FunctionLV
	TotalItems := LV_GetCount()
	a:=15/TotalItems
	loop % LV_GetCount()
	{
		loadval := Round(loadval + a, 2)
		SetLoad(loadval, "Importing Hotkeys (" . A_Index . ")...")
		Sleep, 0
		LV_GetText(Func,A_Index)
		IniRead,TempHotkey,Data.Ini,%Func%,Hotkey
		LV_Modify(A_Index,,,TempHotkey)
	SetHotkeys(Func,TempHotkey,"On")
	}

}

return


LoadNSN()
{
    global NSN,
	global loadval
    NSN := {}
	FileRead, Content, NSNLIST.txt
	RegExReplace(Content, "(\R)",,count)
	a:=20/count
    Loop, Read, NSNLIST.txt
    {
		loadval := Round(loadval + a, 2)
		SetLoad(loadval, "Importing NSN List (" . A_Index . ")...",0)
		;Sleep, -1
        line := Trim(A_LoopReadLine)

        Loop, Parse, line, %A_Tab%
        {
            if (A_LoopField != "")
                NSN[A_LoopField] := 1
        }
    }

    ;MsgBox % "Loaded NSN: " NSN.Count()
}



LoadTable()
{
    xl := ComObjActive("Excel.Application")
    global BatchLP

    xl := ComObjActive("Excel.Application")
    ws := xl.ActiveSheet

    lastRow := ws.Cells(ws.Rows.Count, 3).End(-4162).Row ; Column C

    BatchLP := {}

    Loop % lastRow - 1
    {
        row := A_Index + 1

        Batch := Trim(ws.Cells(row, BatchColumnIndex).Text)
        LP := Trim(ws.Cells(row, LPColumnIndex).Text)

        if (Batch != "")
            BatchLP[Batch] := LP
    }

    ;MsgBox, 64, Load Complete, % "Loaded " BatchLP.Count() " batches"
}

; ===== END OF LOAD DATA FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================



; ===== HOTKEY FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================


F6::
text := "=IF(OR("
counter := 0

; replace ALL separators with space first
clip := A_Clipboard
clip := StrReplace(clip, "`n", " ")
clip := StrReplace(clip, "`r", " ")
clip := StrReplace(clip, ",", " ")
clip := StrReplace(clip, "`t", " ")

Loop, Parse, clip, %A_Space%
{
    temp := Trim(A_LoopField, " `t`r`n,.")

    if (temp = "")
        continue

    if (NSN.HasKey(temp))
    {
        if (counter > 0)
            text .= ","

        text .= "A3=""" temp """"
        counter++
    }
}

text .= "),""NIS"",D3)"

MsgBox % "Matched: " counter
Clipboard := text
return



^!c::
global output
xl := ComObjActive("Excel.Application")
output := ""
prevRow := ""

for cell in xl.Selection
{
    if (prevRow = "")
    {
        output := cell.Value
    }
    else if (cell.Row != prevRow)
    {
        output .= "`r`n" cell.Value
    }
    else
    {
        output .= " " cell.Value
    }

    prevRow := cell.Row
}
tooltip,Copied Cells
clipboard := output
sleep,500
tooltip,
return


f12::   ; Open Calculator
Numlock::

if WinExist("Calculator") {
    WinActivate
    Send, {Delete}
    Send, {BackSpace}
}else
    Run, % "calc.exe"
return



^PgDn::

If (ToggleCount > 0)
	ToggleCount-=20

else
	ToggleCount:=0
WinSet,Transparent,%ToggleCount%,A


return

^PgUp::
If (ToggleCount < 255)
	ToggleCount+=20
else
	ToggleCount:=255
WinSet,Transparent,%ToggleCount%,A
return


f10::
try xl := ComObjActive("Excel.Application")
catch
{
    MsgBox Excel is not open.
	return
}

sheet := xl.ActiveSheet
lastRow := sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row

; Find columns by header
headers := {}

Loop % sheet.UsedRange.Columns.Count
{
    header := Trim(sheet.Cells(1, A_Index).Value)
    headers[header] := A_Index
}

productCol := headers["Product name"]
onOrderCol := headers["On order"]
unitCol := headers["Unit"]

if (!productCol || !onOrderCol || !unitCol)
{
    MsgBox Required columns not found.
    return
}

result := ""

Loop % lastRow - 1
{
    row := A_Index + 1

    onOrder := sheet.Cells(row, onOrderCol).Value

    if (onOrder > 0)
    {
        qty := Round(onOrder)  ; removes decimals
        product := Trim(sheet.Cells(row, productCol).Value)
        unit := Trim(sheet.Cells(row, unitCol).Value)

        result .= qty . unit . " - " . product . "`r`n"
    }
}

Clipboard := RTrim(result, "`r`n")

count := StrSplit(Clipboard, "`n").Length()

MsgBox % "Copied " count " items to clipboard."

return

f11::
global LPColumnIndex
global BatchColumnIndex


Loop
{
    InputBox, BatchColumnIndex, Batch Setup, Enter Batch Column Index: (e.g: 3)

    if ErrorLevel
        return

    BatchColumnIndex := Trim(BatchColumnIndex)

    if RegExMatch(BatchColumnIndex, "^\d+$") && (BatchColumnIndex > 0)
    {
        BatchColumnIndex += 0
        break
    }

    MsgBox, 16, Error, Enter a valid positive number only
}


Loop
{
    InputBox, LPColumnIndex, Column Setup, Enter LP Column Index: (e.g: 6)

    if ErrorLevel
        return

    LPColumnIndex := Trim(LPColumnIndex)

    if RegExMatch(LPColumnIndex, "^\d+$") && (LPColumnIndex > 0)
    {
        LPColumnIndex += 0
        break
    }

    MsgBox, 16, Error, Enter a valid positive number only
}


Tooltip,Please Wait...
LoadTable()
ToolTip,
MsgBox, 64, Load Complete, % "Loaded " BatchLP.Count() " batches"
return


^Esc::
CLose:
Reload
return



^Tab::
modes := ["DeCA 2", "3PL", "SPV"]

index := 1

; find current index
for i, v in modes
{
    if (v = PickingMode)
    {
        index := i
        break
    }
}

ToolTip % "Current: " PickingMode "`nPress TAB to cycle"

Loop
{
    if !GetKeyState("Ctrl", "P")
        break

    if GetKeyState("Tab", "P")
    {
        index++
        if (index > modes.Length())
            index := 1

        PickingMode := modes[index]

        ToolTip % "Mode Selected: " PickingMode

        KeyWait, Tab
    }

    Sleep, 30
}

ToolTip
return




; ===== END OF HOTKEY FUNCTIONS ===================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================
#Include *i Personal Script.ahk

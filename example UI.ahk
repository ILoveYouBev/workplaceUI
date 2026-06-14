#NoEnv
#SingleInstance Force
SetBatchLines, -1

; =========================
; DARK BACKGROUND GUI
; =========================
Gui, Color, 1E1E1E
Gui, Font, s10 cFFFFFF, Segoe UI

; TITLE
Gui, Font, s14 Bold c00AEEF
Gui, Add, Text, x30 y20 w300 Center, Excel Automation Setup

; LABELS
Gui, Font, s10 Bold cFFFFFF

Gui, Add, Text, x30 y70 w200, LP Column Index:
Gui, Add, Edit, x30 y95 w260 vLPInput Center

Gui, Add, Text, x30 y140 w200, Batch Column Index:
Gui, Add, Edit, x30 y165 w260 vBatchInput Center

; BUTTON
Gui, Font, s10 Bold
Gui, Add, Button, x30 y220 w260 h35 gStart, START

; SHOW GUI
Gui, Show, w320 h290, Excel Tool
return

; =========================
; START BUTTON LOGIC
; =========================
Start:
Gui, Submit, NoHide

; VALIDATION (numbers only)
if !RegExMatch(LPInput, "^\d+$") {
    MsgBox, 16, Error, LP Column must be a number
    return
}

if !RegExMatch(BatchInput, "^\d+$") {
    MsgBox, 16, Error, Batch Column must be a number
    return
}

LPColumnIndex := LPInput + 0
BatchColumnIndex := BatchInput + 0

MsgBox, 64, Success,
(
LP Column: %LPColumnIndex%
Batch Column: %BatchColumnIndex%
)

return

; =========================
; CLOSE GUI
; =========================
GuiClose:
ExitApp
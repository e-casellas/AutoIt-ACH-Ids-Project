#include <IE.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "MetroGUI-UDF\MetroGUI_UDF.au3"
#include "MetroGUI-UDF\_GUIDisable.au3"
Opt("WinTitleMatchMode", 2)

;Variables
Global $lLastRow, $oIE, $oExcel, $oWorkbook, $sCompanyId, $sMainCompanyId

;ESC key will stop this bot
HotKeySet("{ESC}", "_Exit")
Func _Exit()
	Exit
EndFunc   ;==>_Exit

#Region GUI
_Metro_EnableHighDPIScaling()
_SetTheme("LightGray")
$GUIThemeColor = 0xeff4ff
$Form1 = _Metro_CreateGUI("ACH IDs Limit Project", 310, 175, -1, -1, True)
$ButtonBKColor = 0x603cba
$Button1 = _Metro_CreateButtonEx2("Start", 210, 50, 80, 30)
$Button2 = _Metro_CreateButtonEx2("Stop", 210, 100, 80, 30)
$Label1 = GUICtrlCreateLabel("ACH IDs Limit Project", 50, 5, 120, 17)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x603cba)
$Label2 = GUICtrlCreateLabel("", 50, 70, 120, 40)
GUICtrlSetFont(-1, 11, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0xee1111)
;GUICtrlSetState(-1, $GUI_HIDE)
$Control_Buttons = _Metro_AddControlButtons(True, False, True, False, True) ;CloseBtn = True, MaximizeBtn = True, MinimizeBtn = True, FullscreenBtn = True, MenuBtn = True
$GUI_CLOSE_BUTTON = $Control_Buttons[0]
$GUI_MINIMIZE_BUTTON = $Control_Buttons[3]
$GUI_MENU_BUTTON = $Control_Buttons[6]
GUISetState(@SW_SHOW)
#EndRegion GUI

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $GUI_CLOSE_BUTTON
			ExitLoop
			Exit
		Case $Form1
		Case $GUI_MINIMIZE_BUTTON
			GUISetState(@SW_MINIMIZE, $Form1)
		Case $GUI_MENU_BUTTON
			Local $MenuButtonsArray[2] = ["About", "Exit"]
			Local $MenuSelect = _Metro_MenuStart($Form1, 50, $MenuButtonsArray)
			Switch $MenuSelect ;Above function returns the index number of the selected button from the provided buttons array.
				Case "0"
				Case "1"
					_Metro_GUIDelete($Form1)
					Exit
			EndSwitch
		Case $Button1
			OpenExcelInput()
			For $i = 2 To $lLastRow Step 1
				RemoveLoop()
				AddLoop()
			Next
			_Excel_Close($oExcel)
		Case $Button2
			ExitLoop
			Exit
		Case $Label1
	EndSwitch
WEnd

Func OpenExcelInput()
	$oExcel = _Excel_Open(False, False, False, False, True)
	$sWorkbook = @ScriptDir & "\input.xlsx"
	$oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
	$lLastRow = $oWorkbook.ActiveSheet.UsedRange.Rows.Count
	$sMainCompanyId = _Excel_RangeRead($oWorkbook, "Sheet1", "B1", 1)
EndFunc   ;==>OpenExcelInput

Func RemoveLoop()
	WinActivate("FIS ACH")
	$oIE = _IEAttach("FIS ACH")
	GUICtrlSetData($Label2, "Working on cell       " & $i & " out of " & $lLastRow)
	GUICtrlSetState($Label2, $GUI_SHOW)
	$sCompanyId = _Excel_RangeRead($oWorkbook, "Sheet1", "A" & $i, 1)
	_Write($oIE, "inp2_searchByCompanyId", $sCompanyId)
	_Click($oIE, "selectButton")
	While 1
		If WinExists("Profile Search Results") Or WinExists("Change Originator") Then
			ExitLoop
			Sleep(100)
		EndIf
	WEnd
	If WinExists("Profile Search Results") Then
		$oIE = _IEAttach("Profile Search Results")
		Sleep(500)
		_Click($oIE, "select")
	EndIf
	Sleep(1000)
	$oIE = _IEAttach("Change Originator")
	Do
		_IEGetObjByName($oIE, "rcorner_menu_contents_LimitsMenu")
	Until Not @error
	_Click($oIE, "rcorner_menu_contents_LimitsMenu")
	Send("{UP}")
	Sleep(100)
	Send("{ENTER}")
	Sleep(100)
	Send("{ENTER}")
	Sleep(2000)
	WinActivate("Change Originator")
	Send("^w")
EndFunc   ;==>RemoveLoop

Func AddLoop()
	WinActivate("FIS ACH")
	$oIE = _IEAttach("FIS ACH")
	_Write($oIE, "inp2_searchByCompanyId", $sCompanyId)
	_Click($oIE, "selectButton")
	While 1
		If WinExists("Profile Search Results") Or WinExists("Originator Setup") Then
			ExitLoop
			Sleep(100)
		EndIf
	WEnd
	If WinExists("Profile Search Results") Then
		$oIE = _IEAttach("Profile Search Results")
		Sleep(500)
		_Click($oIE, "select")
	EndIf
	Sleep(1500)
	$oIE = _IEAttach("Originator Setup")
	_Click($oIE, "LIMITS")
	_Click($oIE, "ExposureLimits")
	_Click($oIE, "DailyLimits")
	For $j = 1 To 24 Step 1
		Send("{TAB}")
	Next
	Send("{SPACE}")
	_Write($oIE, "MainCompanyID", $sMainCompanyId)
	_Click($oIE, "next")
	Do
		_IEGetObjByName($oIE, "BANK_CODE")
	Until Not @error
	_Write($oIE, "BANK_CODE", "PAB DT")
	For $x = 1 To 2 Step 1
		For $y = 1 To 9 Step 1
			;Do
			;_IEGetObjById($oIE, "addRowBtn_LimitsTaskList")
			;Until Not @error
			_Click($oIE, "addRowBtn_LimitsTaskList")
		Next
		_Write($oIE, "Row3_SEC_CODE", "IAT")
		_Write($oIE, "Row4_SEC_CODE", "IAT")
		_Write($oIE, "Row5_SEC_CODE", "TEL")
		_Write($oIE, "Row6_SEC_CODE", "TEL")
		_Write($oIE, "Row7_SEC_CODE", "WEB")
		_Write($oIE, "Row8_SEC_CODE", "WEB")
		_Write($oIE, "Row9_SEC_CODE", "XCK")
		_Write($oIE, "Row10_SEC_CODE", "XCK")

		_Write($oIE, "Row1_DRCR_CODE", "C")
		_Write($oIE, "Row2_DRCR_CODE", "D")
		_Write($oIE, "Row3_DRCR_CODE", "C")
		_Write($oIE, "Row4_DRCR_CODE", "D")
		_Write($oIE, "Row5_DRCR_CODE", "C")
		_Write($oIE, "Row6_DRCR_CODE", "D")
		_Write($oIE, "Row7_DRCR_CODE", "C")
		_Write($oIE, "Row8_DRCR_CODE", "D")
		_Write($oIE, "Row9_DRCR_CODE", "C")
		_Write($oIE, "Row10_DRCR_CODE", "D")

		_Write($oIE, "Row1_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row2_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row3_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row4_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row5_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row6_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row7_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row8_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row9_INFORM_SUSPEND_CODE", "S")
		_Write($oIE, "Row10_INFORM_SUSPEND_CODE", "S")

		_Click($oIE, "next")
	Next
	Sleep(2000)
	;Print()
	WinActivate("Change Originator")
	Send("^w")
EndFunc   ;==>AddLoop

Func Print()
	_IEAction($oIE, "print")
	WinWait("Print")
	WinActivate("Print")
	Sleep(500)
	ControlClick("Print", "", "Button10")
	Sleep(500)
	ControlClick("Print", "", "Button13")
	Sleep(500)
	Do
		WinActivate("Review ACH Profile")
	Until Not @error
	Sleep(500)
	Send("^w")
	Sleep(500)
EndFunc   ;==>Print

#Region MyFunctions ===================================================================
Func _Click($Tab, $ObjIdOrName)
	$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
	If @error Then $oObj = _IEGetObjById($Tab, $ObjIdOrName)
	_IEAction($Obj, "click")
EndFunc   ;==>_Click

Func _Write($Tab, $ObjIdOrName, $Text)
	$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
	If @error Then $oObj = _IEGetObjById($Tab, $ObjIdOrName)
	_IEFormElementSetValue($Obj, $Text)
EndFunc   ;==>_Write
#EndRegion MyFunctions ===================================================================

Exit

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
CoordMode, Mouse, Screen
SetTitleMatchMode RegEx

PDURegEx := "(?:pdu rack (?<Rack>[0-9]{1,5}))|(?:Product S\/N: *(?<Serial>[A-Z0-9-]{1,}))|(?:Model N.*: *(?<Model>[A-Z0-9-]{1,}))|(?:MAC.*: *(?<MAC>(?:[0-9A-Fa-f]{2}[:-]){5}(?:[0-9A-Fa-f]{2})))|(?:(?:(?:\.[A-D]{2} *[A-D]:[X-Z] *(?:On|Off|On\/Fuse) *)|(?:[A-D1-3]{3} *[0-9]{1,2}A.*% *)|(?:\.[A-D]{2} *(?:(?:Master|Link)_[X-Z]|Tower[A-D]_Infeed[A-D]) *(?:On|Off|On\/Fuse) *(?:N\/A|[0-9\.]*) *)|(?:\.[A-B]A *(?:On|Off|On\/Fuse) *))(?<Load>[0-9]{1,2}\.[0-9]{1,2}))|(?:Critical Alert[ -]*(?<CAlert>.*))|(?:pdu end rack (?<EndRack>[0-9]{1,5}))|(?:3-Phase: *(?<Phase>(?:Yes|No)))"
AutoTrim, On

; Define GUI parameters
	;Gui,+AlwaysOnTop
	Gui,Font,Normal s12 c0x0 Bold,Lucida Sans Unicode
	Gui,Add,GroupBox,x10 y10 w200 h130 vGUIDataCollection,Data Collection
	Gui,Font,Normal s14 c0x0,Lucida Console
	Gui,Add,Text,x30 y51 w75 h20 Center vGUIRack,Rack #
	Gui,Font,Normal s12 c0x0,Lucida Console
	Gui,Add,Edit,x110 y50 w75 h21 Number vPDURack,
	Gui,Font,Normal s12 c0x0,Lucida Sans Unicode
	Gui,Add,Button,x55 y90 w105 h30 0xC00 0x1 vGUICollectData gCollectData,Collect Data
	Gui,Font,Normal s12 c0x0 Bold,Lucida Sans Unicode
	Gui,Add,GroupBox,x220 y10 w200 h130 vGUIDataParsing,Data Parsing
	Gui,Font
	Gui,Add,Text,x230 y40 w185 h55 Center vGUIDataParsingText,When ready to parse, click Parse Data button below and use the file selection dialog to choose your log file.
	Gui,Font,Normal s12 c0x0,Lucida Sans Unicode
	Gui,Add,Button,x265 y90 w105 h30 0xC00 vGUIParseData gParseData,Parse Data
	Gui,Show,x519 y308 w430 h160 ,PDU Data Parse Tool v0.8
Return

CollectData:
	GUI, Submit
	If PDURack !> 0
	{
		MsgBox Rack # cannot be blank!
		GUI, Show
		Return
	}
	If WinExist("COM.{1,2} - PuTTY")
	{
		WinActivate
		SendInput {Enter}
		Sleep, 200
		SendInput {Enter}
		Sleep, 200
		SendInput {Enter}
		Sleep, 200
		SendInput admn
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput admn
		Sleep, 100
		SendInput {Enter}
		Sleep, 200
		SendInput pdu rack %PDURack%
		Sleep, 100
		SendInput {Enter}
		MsgBox , , Input Wait, Waiting 5 seconds before inputting commands..., 5
		SendInput show towers
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput show system
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput istat
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput n
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput show units
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput show network
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput n
		Sleep, 100
		SendInput {Enter}
		Sleep, 100
		SendInput lstat
		Sleep, 100
		SendInput {Enter}
		Sleep, 1000
		SendInput pdu end rack %PDURack%
		Sleep, 100
		SendInput {Enter}
		GUI, Show
	}
	Else
	{	
		MsgBox PuTTY not found!
		Gui, Show
		Return
	}
Return

ParseData:
	FileSelectFile, PDUFiles, M3, , Select PDU Logs to parse. Output will be in .\Parsed\
	If PDUFiles =
		Return
	Loop, Parse, PDUFiles, `n
	{
		If A_Index = 1
		{
			PDUDir = %A_LoopField%
			SetWorkingDir, %PDUDir%
			FileCreateDir, %PDUDir%\Parsed
			PDULoadCount = 0
		}
		Else
		{
			PDUFileIn = %A_LoopField%
			PDUFileOut := SubStr(PDUFileIn, 1, -4)
			PDUFileOut = %PDUDir%\Parsed\%PDUFileOut%_parsed.csv
			Loop, Read, %PDUFileIn%, %PDUFileOut%
			{
				PDUData = %A_LoopReadLine%
				If A_Index = 1
					FileAppend, Rack`,Serial`,Model`,MAC`,Load X`,Load Y`,Load Z`,Critical Alert`n
				If RegExMatch(PDUData, PDURegEx, PDUMatch)
				{
					If PDUMatchRack >= 0
					{
						Row1Rack = %PDUMatchRack%
						;Row1Row := SubStr(Row1RackInput, 1, (StrLen(Row1RackInput)-2))
						;Row1Rack := SubStr(Row1RackInput, -1)
					}
					If PDUMatchSerial >= 0
					{
						If Row1Serial >= 0
							Row2Serial = %PDUMatchSerial%
						Else
							Row1Serial = %PDUMatchSerial%
					}
					If PDUMatchModel >= 0
					{
						If Row1Model >= 0
							Row2Model = %PDUMatchModel%
						Else
							Row1Model = %PDUMatchModel%
					}
					If PDUMatchMAC >= 0
					{
						If Row1MAC >= 0
							Row2MAC = %PDUMatchMAC%
						Else
							Row1MAC = %PDUMatchMAC%
					}
					If PDUMatchPhase >= 0
					{
						If PDUMatchPhase = No
						{
							PDUPhase = 1
						}
						Else
						{
							PDUPhase = 3
						}
					}
					If PDUMatchLoad >= 0
					{
						If PDUPhase = 1
						{
							;MsgBox 1-Phase, Count = %PDULoadCount%
							If PDULoadCount <> 2
							{	
								PDULoadCount := PDULoadCount + 1
							}
							Else
								PDULoadCount = 1
							If PDULoadCount = 1
								Row1LoadX = %PDUMatchLoad%
							If PDULoadCount = 2
								Row2LoadX = %PDUMatchLoad%
						}
						Else
						{
							If PDULoadCount <> 3
							{	
								PDULoadCount := PDULoadCount + 1
							}
							Else
								PDULoadCount = 1
							If PDULoadCount = 1
							{
								If Row1LoadX >= 0
									Row2LoadX = %PDUMatchLoad%
								Else
									Row1LoadX = %PDUMatchLoad%
							}
							If PDULoadCount = 2
							{
								If Row1LoadY >= 0
									Row2LoadY = %PDUMatchLoad%
								Else
									Row1LoadY = %PDUMatchLoad%
							}
							If PDULoadCount = 3
							{
								If Row1LoadZ >= 0
									Row2LoadZ = %PDUMatchLoad%
								Else
									Row1LoadZ = %PDUMatchLoad%
							}
						}
					}
					If PDUMatchCAlert >= 0
						RackCAlert = "CRITICAL ALERT"
					If PDUMatchEndRack >= 0
					{
						FileAppend, %Row1Rack%`,%Row1Serial%`,%Row1Model%`,%Row1MAC%`,%Row1LoadX%`,%Row1LoadY%`,%Row1LoadZ%`,%RackCAlert%`n%Row1Rack%`,%Row2Serial%`,%Row2Model%`,%Row2MAC%`,%Row2LoadX%`,%Row2LoadY%`,%Row2LoadZ%`,%RackCAlert%`n
						Row1Rack = ""
						Row2Rack = ""
						Row1Serial = ""
						Row2Serial = ""
						Row1Model = ""
						Row2Model = ""
						Row1MAC = ""
						Row2MAC = ""
						Row1LoadX = ""
						Row1LoadY = ""
						Row1LoadZ = ""
						Row2LoadX = ""
						Row2LoadY = ""
						Row2LoadZ = ""
						PDUPhase = ""
						PDULoadCount = 0
						RackCAlert = ""
					}
				}
			}
		}
	}
Return

GuiClose:
	ExitApp
Return
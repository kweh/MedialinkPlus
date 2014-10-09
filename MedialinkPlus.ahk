SetTitleMatchMode,2
DetectHiddenText, On
#include cxuser.ahk
#persistent
/* 
 
:------------------------------------------------:
:  + ATT GÖRA									 :
:------------------------------------------------:

Automatisk kontroll av flik vid automatinbokning.

*/
; -----------------------------------------
; --------------- INIT --------------------
; -----------------------------------------
version = 1.785

nyheterText =
(
+ Automatinbokningen visar nu felmeddelande om startdatumet har passerat och justerar detta automagiskt.
)

UpdateTip = 0





Gosub, anvNamn

; -----------------------------------------
; --------------- BOOT --------------------
; -----------------------------------------

if !booted

	mlpDir = %A_AppData%\MLP
	mlpSettings = %mlpDir%\settings.ini
	if fileExist(mlpSettings)
	{
		
	} else {
		FileCreateDir, %mlpDir%
		FileAppend,,%mlpDir%\settings.ini
		IniWrite, %version%, %mlpDir%\settings.ini, Version, Version
	}

	; SPLASH
	SetTimer, versionTimer, 5000
	SplashImage, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\mlp1_7.jpg, B
	Sleep 3000
	SplashImage, Off

	; /SPLASH

	; Läs version i ini
	IniRead, ownVersion, %mlpDir%\settings.ini, Version, Version, 0

	; om versionen i ini-filen är mindre än programmets, uppdatera ini
	if (ownVersion < version)
	{
		IniWrite, %version%, %mlpDir%\settings.ini, Version, Version
		IniRead, ownVersion, %mlpDir%\settings.ini, Version, Version, 0
	}

	; Jämför master med ini
	IniRead, masterVersion, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\master.ini, Version, Version, 0
	versionCheck := masterVersion - ownVersion
	if (versionCheck > 0)
	{
		MsgBox,4, Ny Version!, Det finns en ny version av MedialinkPlus!`nVill du hämta den?
		IfMsgBox Yes
    		run, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus
    	IfMsgBox No
    		return
	}

	;Visa nyhetsmeddelande om detta inte gjorts tidigare
	IniRead, iniNyheter, %mlpDir%\settings.ini, Toggles, Nyheter, 0
	if (iniNyheter = version)
	{
		Return
	} else {
		MsgBox, 0, Vad har hänt?, %nyheterText%
		IniWrite, %version%, %mlpDir%\settings.ini, Toggles, Nyheter
	}

	booted = True
	

; -----------------------------------------
; --------------- MENY --------------------
; -----------------------------------------

~LButton::
	if(A_PriorHotkey = A_ThisHotkey && A_TimeSincePriorHotkey < 300)
	{
		MouseGetPos, , , id, control
		IfInString, control, SysListView
		{
			SLV = 1
		} else {
			SLV = 0
		}
		If ((WinActive("ahk_class wxWindowClassNR") and SLV = 1) or WinActive("ahk_class AutoHotkey"))
		{
			menu, lMenu, add, Klar, Klar
			menu, lMenu, Icon, Klar, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\klar.ico
			menu, lMenu, add, Bearbetas, TilldelaBearbeta
			menu, lMenu, Icon, Bearbetas, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\bearbetas.ico
			menu, lMenu, add, Korrektur Skickat, KorrSkickat
			menu, lMenu, Icon, Korrektur Skickat, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\korr.ico
			menu, lMenu, add, Korrektur Klart, KorrKlart
			menu, lMenu, Icon, Korrektur Klart, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\korr.ico
			menu, lMenu, add, Vilande, Vilande
			menu, lMenu, Icon, Vilande, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\vilande.ico
			menu, lMenu, add, Manus på mail, ManusPaMail
			menu, lMenu, Icon, Manus på mail, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\vilande.ico
			menu, lMenu, add, Undersöks, Undersoks
			menu, lMenu, Icon, Undersöks, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\undersoks.ico
			menu, lMenu, add, Arkiverad, Arkiverad
			menu, lMenu, Icon, Arkiverad, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\arkiverad.ico
			menu, lMenu, add, Ny, Ny
			menu, lMenu, Icon, Ny, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\ny.ico
			menu, lMenu, add, Repetition, Repetition
			menu, lMenu, Icon, Repetition, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\rep.ico
			menu, lMenu, add, Lev. Färdig, LevFardig
			menu, lMenu, Icon, Lev. Färdig, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\fardig.ico
			menu, lMenu, show
		}
	}
	return

~RButton::
	MouseGetPos, , , id, control
	IfInString, control, SysListView
	{
		SLV = 1
	} else {
		SLV = 0
	}
	If ((WinActive("ahk_class wxWindowClassNR") and SLV = 1))
	{
	    Click, right
		Send, {Esc}
		Gosub, printCheck ; Kollar om print finns - sätter 'printExist'
		Gosub, KundOrder  ; Hämtar kundnamn och ordernummer och placerar i 'KundOchOrdernummer'	
		Gosub, anvNamn
		menu, context, add, &Hitta Print-PDF, printGet		; Lägger till "Hitta Print-PDF" i menyn
		menu, context, Icon, &Hitta Print-PDF, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\check.ico
		if (printExist = 0)
		{
			menu, context, disable, &Hitta Print-PDF			; Om ingen print finns, avaktivera menyvalet
			menu, context, Icon, &Hitta Print-PDF, c:\Windows\system32\SHELL32.dll, 132
		} 
		else if (printExist = 1) 
		{
			menu, context, enable, &Hitta Print-PDF			; Om print finns, aktivera menyvalet
		}
		menu, context, add, Sök på ordernummer, SokOrder	; Söker på ordernummret
		menu, context, Icon, Sök på ordernummer, c:\Windows\system32\SHELL32.dll, 210
		menu, context, add, Kopiera kundnamn och ordernummer, KopieraOrdernummerKundnamn		; Kopierar kundnamn och ordernummer till clipboard
		menu, context, Icon, Kopiera kundnamn och ordernummer, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\kopiera.ico
		menu, context, add, Boka annons i C&xense, BokaCX	
		menu, context, Icon, Boka annons i C&xense, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\cx.ico
		menu, context, add, Starta Annons (Photoshop), StartaAnnonsPS		; Startar photoshop med mall baserad på annonsstorlek
		menu, context, Icon, Starta Annons (Photoshop), G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\photoshop.ico
		menu, context, add, Starta Annons (Flash), StartaAnnonsFlash		; Startar flash med mall baserad på annonsstorlek
		menu, context, Icon, Starta Annons (Flash), G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\flash.ico
		menu, context, add, &Tilldela och bearbeta, TilldelaBearbeta			; Markerar som "bearbetas" och tilldelar till användaren
		menu, context, Icon, &Tilldela och Bearbeta, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\tilldela.ico
		menu, context, add, Öppna kundmapp, OppnaKundmapp
	    ;menu, context, add, Skapa egen notering, skapaNotering

		; SUBMENY 'STATUS'
		menu, status, add, Klar, Klar
		menu, status, Icon, Klar, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\klar.ico
		menu, status, add, Bearbetas, Bearbetas
		menu, status, Icon, Bearbetas, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\bearbetas.ico
		menu, status, add, Korrektur Skickat, KorrSkickat
		menu, status, Icon, Korrektur Skickat, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\korr.ico
		menu, status, add, Korrektur Klart, KorrKlart
		menu, status, Icon, Korrektur Klart, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\korr.ico
		menu, status, add, Vilande, Vilande
		menu, status, Icon, Vilande, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\vilande.ico
		menu, status, add, Manus på mail, ManusPaMail
		menu, status, Icon, Manus på mail, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\vilande.ico
		menu, status, add, Undersöks, Undersoks
		menu, status, Icon, Undersöks, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\undersoks.ico
		menu, status, add, Arkiverad, Arkiverad
		menu, status, Icon, Arkiverad, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\arkiverad.ico
		menu, status, add, Ny, Ny
		menu, status, Icon, Ny, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\ny.ico
		menu, status, add, Repetition, Repetition
		menu, status, Icon, Repetition, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\rep.ico
		menu, status, add, Lev. Färdig, LevFardig
		menu, status, Icon, Lev. Färdig, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\lmenu\fardig.ico
		menu, status, add, Tilldela ingen, TilldelaIngen
		menu, status, add, Tilldela..., Tilldela
		menu, context, add, Statusar, :Status
		menu, context, Icon, Statusar, c:\Windows\system32\SHELL32.dll, 099
		;menu, context, add,
		menu, context, add, Skicka Korrmail (Ansv. säljare), Korrmail
		menu, context, Icon, Skicka Korrmail (Ansv. säljare), c:\Windows\system32\SHELL32.dll, 157
		menu, Traffic, add, Felaktig order, FelaktigOrder
		menu, Traffic, Icon, Felaktig order, c:\Windows\system32\SHELL32.dll, 110
		menu, Traffic, add, Öppna i AdBooker, OppnaAdBooker
		menu, Traffic, add, Kontrollera printannonser, printBatch
		menu, Traffic, add, Räkna exponeringar, raknaExponeringar
		menu, Traffic, add, Räkna kampanjer, raknaValda
		menu, Traffic, add, Manus på mail, ManusPaMail
		menu, Traffic, add, Undersöks, Undersoks
		menu, context, add, Traffic, :Traffic
		menu, context, Icon, Traffic, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\ico\traffic.ico
		Menu, context, Show						; Visar menyn
		return	
	} else {

	}
return


; -----------------------------------------
; --------------- INFO --------------------
; -----------------------------------------

;~LButton::
;	control =
;	MouseGetPos, , , id, control
;	if (InStr(control, SysListView))
;	{
;	Send ^c
;		sleep, 100
;		filePath = 
;		txtPath = 
;		ControlSetText, Static29, , Atex MediaLink
;		ControlSetText, Static28, , Atex MediaLink
;		Gosub, KundOrder

;		filePath = G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\assets\noteringar\
;		txtPath = %filePath%%OrderNummer%.txt
;		if FileExist(txtPath)
;		{
;		FileRead, NoteringsText, %filePath%%OrderNummer%.txt
;		ControlMove, Static23,,,, 15, Atex MediaLink
;		ControlMove, Static22,,,, 15, Atex MediaLink
;		;ControlMove, Static28,,,, 70, Atex MediaLink
;		ControlMove, Static29,,,, 70, Atex MediaLink
;		ControlSetText, Static28, Notering, Atex MediaLink
;		ControlSetText, Static29, %NoteringsText%, Atex MediaLink
;		return
;		} else {
;		ControlSetText, Static29, , Atex MediaLink
;		ControlSetText, Static28, , Atex MediaLink
;		return
;		}
;	} 
;	if (InStr(control, Button10))
;	{
;		filePath = G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\assets\noteringar\
;		txtPath = %filePath%%OrderNummer%.txt

;		WinWaitActive, Annonsnummer
;		ControlGetText, internaNoteringar, Edit1, Annonsnummer
;		if FileExist(txtPath)
;		{
;			Sleep, 100
;			FileRead, egnaNoteringar, %txtPath%
;			allaNoteringar = %internaNoteringar%`n`r`n`r* * * * * * `n`rEgna noteringar:`n`r%egnaNoteringar%
;			ControlSetText, Edit1, %allaNoteringar%, Annonsnummer
;		}

;	}
;	return




;skapaNotering:
;	filePath = G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\assets\noteringar\
;	FileRead, NoteringsText, %filePath%%OrderNummer%.txt
;	InputBox, inputNotering, Egna noteringar, Ange eller ändra egna noteringar här.,,,,,,,,%NoteringsText%
;	if ErrorLevel
;		Return
;	Else
;	if fileExist(filePath)
;	{
;		FileDelete, %filePath%%OrderNummer%.txt
;		FileAppend, %inputNotering%, %filePath%%OrderNummer%.txt
;		FileRead, NoteringsText, %filePath%%OrderNummer%.txt
;		ControlSetText, Static29, %NoteringsText%, Atex MediaLink
;		} else {
;			FileCreateDir, %filePath%
;			FileDelete, %filePath%%OrderNummer%.txt
;			FileAppend, %inputNotering%, %filePath%%OrderNummer%.txt
;			FileRead, NoteringsText, %filePath%%OrderNummer%.txt
;			ControlSetText, Static29, %NoteringsText%, Atex MediaLink
;			return
;		}

;		return


; -----------------------------------------
; --------------- MENYVAL -----------------
; -----------------------------------------

anvNamn:
	WinGetTitle, Windowtext, Atex MediaLink
	StringSplit, WindowSplit, Windowtext, =
	Anvandare =  %WindowSplit2%
	return


printCheck:
	Send, ^c
	;Sleep, 50
	SokVag = \\nt.se\Adbase\Annonser\Ad\
	OrderNummer := Clipboard
	StringTrimRight, OrderNummerUtanMnr, OrderNummer, 3
	StringTrimLeft, SistaTvaSiffrorna, OrderNummerUtanMnr, 8
	StringTrimLeft, OrderNummerUtanNollor, OrderNummerUtanMnr, 3
	NySokVag = %SokVag%%SistaTvaSiffrorna%\10%OrderNummerUtanNollor%-01.pdf
	printExist = 0
	IfExist, %NySokVag%
	{
		printExist = 1
	    return
	} else {
		return
	}
	return

KundOrder:
	OrderNummer := Clipboard
	ControlGetText,kund,Static32, Atex MediaLink
	StringReplace, kund, kund, Kontaktperson, , All
	StringReplace, kund, kund, `n, , All
	StringReplace, kund, kund, `r, , All
	KundOchOrdernummer = %kund% (%OrderNummer%)
	return

printGet:
	Run, %nySokVag%
	return

SokOrder:
	ControlFocus, Edit1, A
	Send, {Backspace 15}%OrderNummer%{Backspace 3}{Enter}
	return

Korrmail:
	ControlGetText,epost,Static54, Atex MediaLink
	Send, !s{TAB}akkkkkk{TAB}{Enter}
	Sleep, 50
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% mailade korrektur på %OrderNummer% till %epost%`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	WinActivate, Microsoft Outlook
	Sleep, 100
	Send, ^n
	Sleep, 50
	Send, %epost%{Alt down}m{Alt up} Webbkorr: %KundOchOrdernummer%{Tab}
	return

FelaktigOrder:
	Sleep, 50
	Send, !a{TAB}%Anvandare%{TAB}{Enter}
	Sleep, 50
	Send, !s{TAB}au{TAB}{Enter}

	ControlGetText,epost,Static54, Atex MediaLink
	Gui, Add, Text, x12 y10 w180 h20 , Ange vilka fel som finns på ordern:
	Gui, Add, CheckBox, x12 y40 w180 h30 vGUIsen , Sent bokad
	Gui, Add, CheckBox, x12 y70 w180 h30 vGUImanus, Manus saknas/ej komplett
	Gui, Add, CheckBox, x12 y100 w180 h30 vGUImaterial, Material saknas
	Gui, Add, Button, x52 y130 w90 h30 gOK , OK
	Gui, Show, xCenter yCenter h166 w208, New GUI Window
	return

KopieraOrdernummerKundnamn:
	IfInString, kund, &&
	{
		StringReplace, kund, kund, &&, &
	}
	KundOrder = %kund% (%OrderNummer%)
	Clipboard = %KundOrder%
	ToolTip, Kopierat: %KundOrder%
		SetTimer, RemoveToolTip, 1500
	return

OppnaAdbooker:
	Sleep, 50
	IfWinNotExist, AdBooker
	{
		MsgBox,0, Fel, AdBooker måste startas för att denna funktion ska fungera.
		return
	}
	WinActivate, AdBooker
	Send, ^k
	Send, {Tab 13}
	Send, %OrderNummer%
	Send, ^o
	return


; -----------------------------------------
; --------------- STATUSAR ----------------
; -----------------------------------------
TilldelaBearbeta:
	Sleep, 50
	Send, !a{TAB}%Anvandare%{TAB}{Enter}
	Send, !s{TAB}ab{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Bearbetas"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Klar:
	Sleep, 50
	Send, !s{TAB}ak{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Klar"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Bearbetas:
	Sleep, 50
	Send, !s{TAB}ab{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Bearbetas"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

KorrSkickat:
	Sleep, 50
	Send, !s{TAB}akkkkkk{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Korrektur Skickat"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return


KorrKlart:
	Sleep, 50
	Send, !s{TAB}akkkkk{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Korrektur Klart"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Vilande:
	Sleep, 50
	Send, !s{TAB}av{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Vilande"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

ManusPaMail:
	Sleep, 50
	Send, !a{TAB}mmmm{TAB}{Enter}
	Send, !s{TAB}av{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Manus på Mail"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Undersoks:
	Sleep, 50
	Send, !a{TAB}%Anvandare%{TAB}{Enter}
	Send, !s{TAB}au{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Undersöks"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Arkiverad:
	Sleep, 50
	Send, !s{TAB}a{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Arkiverad"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Ny:
	Sleep, 50
	Send, !s{TAB}n{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Ny"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

Repetition:
	Sleep, 50
	Send, !s{TAB}rr{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Repetition"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

LevFardig:
	Sleep, 50
	Send, !s{TAB}l{TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% satte %OrderNummer% till "Lev. Färdig"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return

FlerStatusar:
	Sleep, 50
	Send, !s
	Return

Tilldela:
	Sleep, 50
	Send, !a
	return

TilldelaIngen:
	Sleep, 50
	Send, !a{TAB} {TAB}{Enter}
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% tog bort tilldelning på %OrderNummer% `n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	return


; -----------------------------------------
; --------------- GUI ---------------------
; -----------------------------------------

GuiClose:
	Gui, Destroy
	return


OK:
	Gui, Submit
	Gui, Destroy
	if GUIsen = 1
	{
		Sentbokad = Sent bokad (material och manus ska finnas 2 arbetsdagar innan startdatum). Om inte ordern bokas fram finns risk att den inte blir producerad.`n
	} else {
		Sentbokad =
	}

	if GUImanus = 1
	{
		Manus = Manus saknas/är inkomplett i interna noteringar`n
	} else {
		Manus =
	}

	if GUImaterial = 1
	{
		Material = Material saknas eller har ett startdatum senare än webbannonsens. Material ska vara oss tillhanda 2 arbetsdagar innan annonsstart.`n
	} else {
		Material =
	}
	FileAppend,
		(
			%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% mailade %epost% ang. felaktig order (%OrderNummer%) och satte den som  "Undersöks"`n
		),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	WinActivate, Microsoft Outlook
	Sleep, 100
	Send, ^n
	Sleep, 50
	Send, %epost%{Alt down}m{Alt up} Ej komplett/korrekt order: %KundOchOrdernummer%{Tab}
	Sleep, 50
	SetKeyDelay, 0
	Send, Din bokning, {Ctrl down}f{Ctrl up}%KundOchOrdernummer%{Ctrl down}f{Ctrl up}, är inte komplett/korrekt och kan inte produceras till det datum den är bokad.`nVar vänlig se över bokningen.`n`n{Ctrl down}f{Ctrl up}Upptäckt problem:{Ctrl down}f{Ctrl up}`n%Sentbokad%%Manus%%Material%
	SetKeyDelay, 10
	return

; -----------------------------------------
; --------- STARTA ANNONS -----------------
; -----------------------------------------

StartaAnnonsPS:
	Send, !a{TAB}%Anvandare%{TAB}{Enter}
	Send, !s{TAB}ab{TAB}{Enter}
	Sleep, 50
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% startade produktion av %OrderNummer% i Photoshop. Status "Bearbetas"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	Gosub, getFormat
	Gosub, getTidning
	ControlGetText, datum, Static13, Atex MediaLink
	StringReplace, datum, datum,-,,All
	StringReplace, kund, kund,:,,All
	StringReplace, kund, kund,\,,All
	StringReplace, kund, kund,/,,All
	StringTrimLeft,datum,datum,2
	StringLen, tecken, kund
	tecken := tecken - 1
	StringTrimRight, forstaBokstavKundnamn, kund, %tecken%

	filePath = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstavKundnamn%\%kund%\%datum%\
	Template = G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\assets\copy\
	if fileExist(filePath)
	{
		FileCopy, %Template%%format%.psd, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.psd
		run, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.psd
	} else {
		FileCreateDir, %filePath%
		FileCopy, %Template%%format%.psd, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.psd
		run, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.psd
	}
	return

StartaAnnonsFlash:
	Send, !a{TAB}%Anvandare%{TAB}{Enter}
	Send, !s{TAB}ab{TAB}{Enter}
	Sleep, 50
	FileAppend,
	(
		%A_YYYY%-%A_MM%-%A_DD%  %A_Hour%:%A_Min%   %Anvandare% startade produktion av %OrderNummer% i Flash. Status "Bearbetas"`n
	),G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\log\log.txt
	Gosub, getFormat
	Gosub, getTidning
	ControlGetText, datum, Static13, Atex MediaLink
	StringReplace, datum, datum,-,,All
	StringReplace, kund, kund,:,,All
	StringReplace, kund, kund,\,,All
	StringReplace, kund, kund,/,,All
	StringTrimLeft,datum,datum,2
	StringLen, tecken, kund
	tecken := tecken - 1
	StringTrimRight, forstaBokstavKundnamn, kund, %tecken%
	filePath = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstavKundnamn%\%kund%\%datum%\
	Template = G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\assets\copy\
	if fileExist(filePath)
	{
		FileCopy, %Template%%format%.fla, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.fla
		run, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.fla
	} else {
		FileCreateDir, %filePath%
		FileCopy, %Template%%format%.fla, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.fla
		run, %filePath%%mlTidning%%mlFormat%-%kund%-%datum%.fla
	}
	return


OppnaKundmapp:
	ControlGetText, datum, Static13, Atex MediaLink
	StringReplace, datum, datum,-,,All
	StringTrimLeft,datum,datum,2
	StringLen, tecken, kund
	tecken := tecken - 1
	StringReplace, kund, kund,:,,All
	StringReplace, kund, kund,\,,All
	StringReplace, kund, kund,/,,All
	StringTrimRight, forstaBokstavKundnamn, kund, %tecken%
	filePath = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstavKundnamn%\%kund%\
	if fileExist(filePath)
	{
		run, %filePath%
	} else {
		MsgBox, Mappen "%filePath% finns inte.
	}

	return

; -----------------------------------------
; --------------- SUBS --------------------
; -----------------------------------------


getKundnamn:
	ControlGetText, mlKundnamn, Static32, Atex MediaLink
	IfInString, mlKundnamn, &&
	{
		StringReplace, mlKundnamn, mlKundnamn, &&, &
	}
	Return

getTidning:
	ControlGetText, tidning, Static23, Atex MediaLink
	StringSplit, tidningRaw, tidning, %A_Space%
	mlTidning = %tidningRaw1%
	if mlTidning = NT
	{
		mlTidning = NTFB
	}
	return

getFormat:
	ControlGetText, format, Static19, Atex MediaLink	
	if (format = "320 x 80" or format = "320 x 160" or format = "320 x 320")
	{
		mlFormat = MOB
	}

	if (format = "250 x 240" or format = "200 x 240" or format = "250 x 360")
	{
		mlFormat = WID
	}

	if (format = "980 x 120" or format = "980 x 240" or format = "980 x 480")
	{
		mlFormat = PAN
	}

	if (format = "468 x 240" or format = "468 x 480")
	{
		mlFormat = MOD
	}

	if (format = "468 x 360" or format = "468 x 480")
	{
		mlFormat = MOD
	}

	if (format = "250 x 600" or format = "200 x 600")
	{
		mlFormat = OUT
	}

	if (format = "1920 x 1080")
	{
		mlFormat = 1920
	}

	if (format = "1080 x 1920")
	{
		mlFormat = 1080
	}

	if (format = "1024 x 384")
	{
		mlFormat = 1024
	}

	if (format = "768 x 384")
	{
		mlFormat = 768
	}

	if (format = "512 x 128")
	{
		mlFormat = 512
	}
	return


RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
return


printBatch:
StatusBarGetText, xAds, 3, Atex MediaLink
StringSplit, xAds, xAds, %A_Space%
xAds = %xAds1%
msgbox, Vill du gå igenom %xAds% annonser?
adCount = 0
printAds = 0
while (adCount < xAds)
{
	adCount++
	Send, {Control Down}
	Sleep, 50
	Send, C
	Sleep, 50
	Send, {Control Up}
	SokVag = \\nt.se\Adbase\Annonser\Ad\
	OrderNummer := Clipboard
	StringTrimRight, OrderNummerUtanMnr, OrderNummer, 3
	StringTrimLeft, SistaTvaSiffrorna, OrderNummerUtanMnr, 8
	StringTrimLeft, OrderNummerUtanNollor, OrderNummerUtanMnr, 3
	NySokVag = %SokVag%%SistaTvaSiffrorna%\10%OrderNummerUtanNollor%-01.pdf
	sleep, 50
	IfExist, %NySokVag%
	{
		Send, !s{TAB}av{TAB}{Enter}
		printAds++
		sleep, 50

	}
	sleep, 50
	Send, {Down}
}
msgbox, Klar!`n%printAds% annonser av %xAds% hade printmaterial tillgängligt.
return

raknaExponeringar:
	send, {esc}
	Gosub, anvNamn
	sleep, 150
	ControlGet, exponeringar, List, Selected, SysListView321, Atex MediaLink
	sleep, 100
	ControlGet, antal, List, Count Selected, SysListView321, Atex MediaLink
	totExp = 0
	count = 0
	if (Anvandare = "martinve")
	{
		Loop, Parse, exponeringar, `n
		{
			StringSplit, expCount, A_LoopField, `t
			totExp := totExp + expCount13

			count++
			if (count = antal)
			{
				break
			}
		}
	} else {
		Loop, Parse, exponeringar, `n
		{
			StringSplit, expCount, A_LoopField, `t
			totExp := totExp + expCount10

			count++
			if (count = antal)
			{
				break
			}
		}
	}
	msgbox, Totalt antal begärda exponeringar: %totExp%
	return

raknaValda:
	ControlGet, antal, List, Count Selected, %control%, Atex MediaLink
	Msgbox, %antal% kampanjer markerade.
	return


; //////////////////////////////////////////////////
; CXENSE INBOKNING
; //////////////////////////////////////////////////

BokaCX:
	FileCreateDir, %A_AppData%\AHK
	userFolder = %A_AppData%\AHK\
	; VARIABLE RESET
	mlStartdatum = 0
	mlStoppdatum = 0
	mlExponeringar = 0
	mlKundnr = 0
	mlKundnamn = 0
	mlOrdernr = 0
	mlTidning = 0
	mlFormat = 0


	Gosub, CxenseBokning
	return

getOrdernr:
	Sleep, 50
	Send, ^C
	Sleep, 50
	mlOrdernr := Clipboard
	Return

getFromList:
	MouseGetPos, , , id, control
	IfInString, control, SysListView
	{
	ControlGet, getList, List, Selected, %control%, Atex MediaLink
	StringSplit, getListRow, getList, `n
	listRow = %getListRow1%
	Stringsplit, listArr, listRow, `t
	mlStartdatum = %listArr9%
	mlStoppdatum = %listArr10%
	mlExponeringar = %listArr14%
	mlKundnr = %listArr15%
	mlKundnamn = %listArr16%
	Gosub, anvNamn
	if (Anvandare = "martinve")
		{
			mlStartdatum = %listArr9%
			mlStoppdatum = %listArr10%
			mlExponeringar = %listArr13%
			mlKundnr = %listArr7%
			mlKundnamn = %listArr8%
		}
	if (Anvandare = "anniejo")
		{
			mlStartdatum = %listArr5%
			mlStoppdatum = %listArr6%
			mlExponeringar = %listArr11%
			mlKundnr = %listArr10%
			mlKundnamn = %listArr8%
		}
	if (Anvandare = "annicasv")
		{
			mlStartdatum = %listArr5%
			mlStoppdatum = %listArr6%
			mlExponeringar = %listArr9%
			mlKundnr = %listArr10%
			mlKundnamn = %listArr11%
		}
	if (Anvandare = "mikaellu")
		{
			mlStartdatum = %listArr8%
			mlStoppdatum = %listArr9%
			mlExponeringar = %listArr12%
			mlKundnr = %listArr13%
			mlKundnamn = %listArr5%
		}
	if (Anvandare = "dennisst")
		{
			mlStartdatum = %listArr6%
			mlStoppdatum = %listArr7%
			mlExponeringar = %listArr10%
			mlKundnr = %listArr11%
			mlKundnamn = %listArr12%
		}

	}
	return

; ---------------------------------------
; ----PRODUKT-IDn -----------------------
; ---------------------------------------

;---- Riktade
RiktadMOD = 00000001609df500
RiktadOUT = 00000001609df509
RiktadPAN = 00000001609df502
RiktadWID = 00000001609df506

;---- ROS
RosMOD = 0000000160209522
RosOUT = 0000000160209517
RosPAN = 00000001601d3817
RosWID = 0000000160209515

;---- PLUGG
PluggMOD = 000000016020bf61
PluggOUT = 000000016020bf38
PluggPAN = 000000016020bf82
PluggWID = 000000016020bf85






CxenseBokning:
	Gosub, getFromList			; mlStartdatum, mlStoppdatum, mlKundnr, mlExponeringar, mlKundnamn
	StringReplace, mlKundnamn, mlKundnamn,:,,All
	StringReplace, mlKundnamn, mlKundnamn,\,,All
	StringReplace, mlKundnamn, mlKundnamn,/,,All
	StringReplace, mlKundnamn, mlKundnamn,&&,,All
	StringReplace, mlKundnamn, mlKundnamn,&,,All
	Gosub, getOrdernr			; mlOrdernr
	Gosub, getTidning			; mlTidning
 	Gosub, getFormat			; mlFormat
	Gosub, xmlGET
	if (mlFormat = 0 || mlKundnamn = 0 || mlOrdernr = 0 || mlTidning = 0)
	{
		MsgBox, 48, Eeeh.. oops., Något har gått snett, prova igen!
		Goto, TheEND
	}
	checkKundNR = -%A_Space%%mlKundnr%%A_Space%-

	FileRead, xmlOut, %userFolder%xmlOut.xml
	StringReplace, xmlOut, xmlOut, <cx:childFolder>, +, A
	xmlPart = 0
	Loop, Parse, xmlOut, +
	{
		xmlIndex++
		if InStr(A_LoopField, checkKundNR)
		{
			xmlPart = %A_LoopField%
			break
		}
	}
	if (xmlPart != 0)
	{
		StringSplit, xmlSplit, xmlPart, >
		StringSplit, xmlSplit, xmlSplit4, <
		xmlID = %xmlSplit1% ; xmlID innehåller kundens ID
	}
	if (xmlPart = 0) ; Kund kunde inte hittas
	{
		MsgBox, 4, Kund saknas, Kund fanns inte. Skapa kund "%mlTidning% - %mlKundnr% - %mlKundnamn%"?
		IfMsgBox Yes 
		{
		FileDelete, %userFolder%bat.bat
		FileDelete, %userFolder%xml.xml
		FileDelete, %userFolder%xmlOut.xml
		FileDelete, %userFolder%create.xml
		xmlToRun =
(
<?xml version="1.0" encoding="UTF-8"?>
<cx:folder xmlns:cx="http://cxense.com/cxad/api/cxad">
  <cx:name>%mlTidning% - %mlKundnr% - %mlKundnamn%</cx:name>
</cx:folder>
)
		batToRun = 
(
G:
cd G:\NTM\NTM Digital Produktion\cURL\bin
curl -s -H "Content-type: text/xml" -u %cxUser% -X POST https://cxad.cxense.com/api/secure/folder/advertising -d @%userFolder%xml.xml > %userFolder%create.xml
pause
)
		FileAppend, %batToRun%, %userFolder%bat.bat
		FileEncoding, UTF-8-RAW
		FileAppend, %xmlToRun%, %userFolder%xml.xml
		FileEncoding
		Sleep, 100
		Run, %userFolder%bat.bat
		Sleep, 150
		WinWaitClose, C:\Windows\system32\cmd.exe
		MsgBox,1, Kund inlagd, Kund inlagd. Boka kampanj?
		ifMsgBox, Cancel
		{
			Goto, TheEND
		}
		ifMsgBox, OK
		{
			FileRead, xmlCreate, %userFolder%create.xml
			StringSplit, xmlSplit, xmlCreate, >
			StringSplit, xmlSplit, xmlSplit6, <
			xmlID = %xmlSplit1% ; xmlID innehåller kundens ID
			Goto, bokaKampanjCX
		}
		return
		}
		IfMsgBox, No
		{
			Goto, TheEND
		}
		Goto, bokaKampanjCX
	return
	}
	Goto, bokaKampanjCX
	Return
return


;bokaCampaignCX:
;	Type = 0
;	Gui, Add, DropDownList, x4 y20 w158 h40 vType R4, Run On Site||Riktad|Plugg|Reach
;	Gui, Add, Button, x42 y50 w100 h30 gBokaKampanj, OK
;	Gui, Show, xCenter yCenter h103 w190 , Välj typ
;	return



xmlGET:
FileDelete, %userFolder%bat.bat
FileDelete, %userFolder%xmlOut.xml
sleep, 100
batToRun = 
(
G:
cd G:\NTM\NTM Digital Produktion\cURL\bin
curl -s -H "Content-type: text/xml" -u %cxUser% -X GET https://cxad.cxense.com/api/secure/folder/advertising > %userFolder%xmlOut.xml
)
FileAppend, %batToRun%, %userFolder%bat.bat
Run, %userFolder%bat.bat
Sleep, 150
WinWaitClose, C:\Windows\system32\cmd.exe
fileToCheck = %userFolder%xmlOut.xml
Loop, 40
{
	if fileExist(fileToCheck)
	{
		break
	}
	Sleep, 250
}
if !fileExist(fileToCheck)
{
	Goto, TheEND
}
return



; ------------------- ;


^!D::
	FileRead, xmlOut, %userFolder%xmlOutDemo.xml
	StringSplit, xmlSplit, xmlOut, >
	StringSplit, xmlSplit, xmlSplit6, <
	campaignID = %xmlSplit1% ; campaignID innehåller kampanjens ID
	MsgBox, ID: %campaignID%


bokaKampanjCX:
	advertisingFolder = %mlTidning% - %mlKundnr% - %mlKundnamn%
	campaign = %mlTidning% - %mlFormat% - %mlOrdernr%
	
	;---- Preset
	Type = 0
	FileDelete, %userFolder%xml.xml
	FileDelete, %userFolder%xmlOut.xml
	Sleep, 500
	FileAppend, %xmlToRun%, %userFolder%xml.xml
	StringReplace, mlStartdatumStrip, mlStartdatum, - ,, All
	StringReplace, mlStoppdatumStrip, mlStoppdatum, - ,, All
	StringTrimLeft, mlStartdatumStripYY, mlStartdatumStrip, 2
	FormatTime, idag, , yyyyMMdd
	checkDate := mlStartdatumStrip - idag
	if (checkDate < 0)
	{
		MsgBox,48,Fel i startdatum, Startdatumet på kampanjen har redan passerat. `r`nDagens datum har satts som startdatum istället.
		mlStartdatumStrip = %idag%
	}

	; -- NY BOKNINGSMENY
	Gui, Add, GroupBox, x12 y10 w420 h180 , Bokningsöversikt
	Gui, Add, Text, x22 y40 w80 h20 , Typ:
	Gui, Add, DropDownList, x142 y40 w120 h20 vType R4, Run On Site||Riktad|Plugg|Reach
	Gui, Add, Text, x22 y70 w100 h20 , Advertising Folder:
	Gui, Add, Edit, x142 y70 w280 h20 vadvertisingFolder, %advertisingFolder%
	Gui, Add, Text, x22 y100 w100 h20 , Campaign:
	Gui, Add, Edit, x142 y100 w280 h20 vcampaign, %campaign%
	Gui, Add, Text, x22 y130 w110 h20 , Start- och stoppdatum:
	Gui, Add, DateTime, x142 y130 w120 h20 vmlStartdatum Choose%mlStartdatumStrip%, yyyy-MM-dd
	Gui, Add, Text, x282 y130 w10 h20 , -
	Gui, Add, DateTime, x302 y130 w120 h20 vmlStoppdatum Choose%mlStoppdatumStrip%, yyyy-MM-dd
	Gui, Add, Text, x22 y160 w100 h20 , Exponeringar:
	Gui, Add, Edit, x142 y160 w280 h20 vmlExponeringar, %mlExponeringar%
	Gui, Add, Button, x332 y200 w100 h30 gAvbryt, Avbryt
	Gui, Add, Button, x222 y200 w100 h30 gBokningOK, OK
	Gui, Show, xCenter yCenter h250 w451, Bokningsöversikt
	Return

BokningOK:
		Gui, Submit
		Gui, Destroy
		Gosub, productGET ; returnerar ProductID
		if (Type = "Plugg")
		{
			campaign = %mlTidning% - %mlFormat% - PLUGG - %mlOrdernr%
		}
		if (Type = "Reach")
		{
			campaign = %mlTidning% - REACH - %mlOrdernr%
		}

		xmlToRun =
	(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:campaign xmlns:cx="http://cxense.com/cxad/api/cxad" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<cx:name>%campaign%</cx:name>
	<cx:productId>%productID%</cx:productId>
	<cx:perUserCap>
	<cx:max>0</cx:max>
	<cx:period>DAILY</cx:period>
	</cx:perUserCap>
</cx:campaign>
	)
	FormatTime, mlStartdatum, %mlStartdatum%, yyyy-MM-dd
	FormatTime, mlStoppdatum, %mlStoppdatum%, yyyy-MM-dd
	Gosub, CPMcheck
	FileDelete, %userFolder%xml.xml
	FileDelete, %userFolder%xmlOut.xml
	Sleep, 500
	FileEncoding, UTF-8-RAW
	FileAppend, %xmlToRun%, %userFolder%xml.xml
	FileEncoding
	FileDelete, %userFolder%bat.bat
	batToRun = 
	(
	G:
	cd G:\NTM\NTM Digital Produktion\cURL\bin
	curl -s -H "Content-type: text/xml" -u %cxUser% -X POST https://cxad.cxense.com/api/secure/campaign/%xmlID% -d @%userFolder%xml.xml > %userFolder%xmlOut.xml
	)
	FileAppend, %batToRun%, %userFolder%bat.bat
	Run, %userFolder%bat.bat
	Sleep, 150
	WinWaitClose, C:\Windows\system32\cmd.exe
	FileRead, xmlOut, %userFolder%xmlOut.xml
	StringSplit, xmlSplit, xmlOut, >
	StringSplit, xmlSplit, xmlSplit6, <
	campaignID = %xmlSplit1% ; campaignID innehåller kampanjens ID
	xmlToRun = 
(
<?xml version="1.0" encoding="utf-8"?>
<cx:cpmContract xmlns:cx="http://cxense.com/cxad/api/cxad">
<cx:startDate>%mlStartdatum%T00:00:00.000+02:00</cx:startDate>
<cx:endDate>%mlStoppdatum%T23:59:59.000+02:00</cx:endDate>
<cx:priority>0.50</cx:priority>
<cx:requiredImpressions>%mlExponeringar%</cx:requiredImpressions>
<cx:costPerThousand class="currency" currencyCode="SEK" value="%CPM%"/>
</cx:cpmContract>
)
		FileDelete, %userFolder%xml.xml
		FileEncoding, UTF-8-RAW
		FileAppend, %xmlToRun%, %userFolder%xml.xml
		FileEncoding
		sleep, 100
		FileDelete, %userFolder%bat.bat
		batToRun = 
		(
		G:
		cd G:\NTM\NTM Digital Produktion\cURL\bin
		curl -s -H "Content-type: text/xml" -u %cxUser% -X POST https://cxad.cxense.com/api/secure/contract/%campaignID% -d @%userFolder%xml.xml > %userFolder%xmlOut.xml
		)
		FileAppend, %batToRun%, %userFolder%bat.bat
		Run, %userFolder%bat.bat
		Sleep, 500
		
		FileDelete, %userFolder%xml.xml
		FileDelete, %userFolder%bat.bat
		xmlToRun = 
(
<?xml version="1.0" encoding="UTF-8"?>
<cx:ad xmlns:cx="http://cxense.com/cxad/api/cxad">
  <cx:name>%mlKundnamn% - %mlStartdatumStripYY%</cx:name>
</cx:ad>
)
		batToRun = 
(
G:
cd G:\NTM\NTM Digital Produktion\cURL\bin
curl -s -H "Content-type: text/xml" -u %cxUser% -X POST https://cxad.cxense.com/api/secure/ad/%campaignID% -d @%userFolder%xml.xml
)
		FileEncoding, UTF-8-RAW
		FileAppend, %xmlToRun%, %userFolder%xml.xml
		FileEncoding
		FileAppend, %batToRun%, %userFolder%bat.bat
		Run, %userFolder%bat.bat
		Sleep, 100
		WinWaitClose, C:\Windows\system32\cmd.exe
		MsgBox,4, Bokning klar, Inbokning klar, öppna i webbläsaren?
		IfMsgBox, Yes
			run, https://cxad.cxense.com/adv/campaign/%campaignID%/overview
		IfMsgBox, No
			return
	return

return


productGET:
	;---- Riktade
	RiktadMOD = 00000001609df500
	RiktadOUT = 00000001609df509
	RiktadPAN = 00000001609df502
	RiktadWID = 00000001609df506

	;---- ROS
	RosMOD = 0000000160209522
	RosOUT = 0000000160209517
	RosPAN = 00000001601d3817
	RosWID = 0000000160209515

	;---- PLUGG
	PluggMOD = 000000016020bf61
	PluggOUT = 000000016020bf38
	PluggPAN = 000000016020bf82
	PluggWID = 000000016020bf85

	;---- REACH
	Reach = 000000015f460c65

	if (mlFormat = "MOD" and Type = "Run On Site")
	{
		productID = %RosMOD%
	} 
	else if (mlFormat = "OUT" and Type = "Run On Site")
	{
		productID = %RosOUT%
	}
	else if (mlFormat = "WID" and Type = "Run On Site")
	{
		productID = %RosWID%
	}
	else if (mlFormat = "PAN" and Type = "Run On Site")
	{
		productID = %RosPAN%
	} 
	else if (mlFormat = "MOD" and Type = "Riktad")
	{
		productID = %RiktadMOD%
	} 
	else if (mlFormat = "OUT" and Type = "Riktad")
	{
		productID = %RiktadOUT%
	}
	else if (mlFormat = "WID" and Type = "Riktad")
	{
		productID = %RiktadWID%
	}
	else if (mlFormat = "PAN" and Type = "Riktad")
	{
		productID = %RiktadPAN%
	}
		else if (mlFormat = "MOD" and Type = "Plugg")
	{
		productID = %PluggMOD%
	} 
	else if (mlFormat = "OUT" and Type = "Plugg")
	{
		productID = %PluggOUT%
	}
	else if (mlFormat = "WID" and Type = "Plugg")
	{
		productID = %PluggWID%
	}
	else if (mlFormat = "PAN" and Type = "Plugg")
	{
		productID = %PluggPAN%
	}
	if (Type = "Reach")
	{
		productID = %Reach%
	}
return

CPMcheck:
if (Type = "Plugg")
{
	CPM = 50.00
} 
else if (Type = "Riktad")
{
	CPM = 150.00
}
else if (Type = "Run On Site")
{
	CPM = 100.00
}
else if (Type = "Reach")
{
	CPM = 100.00
}
return

Avbryt:
	Gui, Destroy
	return

TheEND:
	return


versionTimer:
IniRead, masterVersion, G:\NTM\NTM Digital Produktion\Övrigt\MedialinkPlus\dev\master.ini, Version, Version, 0
	versionCheck := masterVersion - ownVersion
	if (versionCheck > 0 and UpdateTip = 0)
	{
		UpdateTip = 1
		TrayTip, Ny version finns!, Det finns en ny version av MedialinkPlus tillgänglig!
	}
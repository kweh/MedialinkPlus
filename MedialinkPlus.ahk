SetTitleMatchMode, 2
DetectHiddenText, On
#SingleInstance force


/*
   ___________________________________________________________
  |                                                           |
  |                         ATT GÖRA                          |
  |___________________________________________________________|
   \_________________________________________________________/
    |::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    |
    | • Kontroll av kampanjID innan inbokning
    |
    |
    |
    |
    |
    |
    |::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯




*/


/*
-----------------------------------------------------------
	Miljövariabler
-----------------------------------------------------------
*/

#include cxuser.ahk
#include httpreq.ahk

version = 2.422
menuOn = 0
lmenuOn = 0
toolbar = 0
traytip = 0
noteWin = 1
skin = 
nyttcitat = 0
citat := 
author =
weblinkget = 0

mlpDir = G:\NTM\NTM Digital Produktion\MedialinkPlus\user\%A_UserName% ; Sätter användarens användarmapp
ifNotExist, %mlpDir% ; om mappen inte finns
	FileCreateDir, %mlpDir% ; skapa mappen

cxDir = %A_AppData%\MLP\cx
IfNotExist, %cxDir%
	FileCreateDir, %cxDir%
IfNotExist, %mlpDir%\skin
	FileCreateDir, %mlpDir%\skin


iconDir = G:\NTM\NTM Digital Produktion\MedialinkPlus\dev\ico ; Sätter mapp för ikoner
templateDir = G:\NTM\NTM Digital Produktion\MedialinkPlus\assets\toCopy ; Sätter mapp för psd- pch fla-mallar
webbannonsDir = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%
lagerDir = X:\digital.ntm.eu\lager
weblinkDir = X:\digital.ntm.eu\weblink
notesDir = G:\NTM\NTM Digital Produktion\MedialinkPlus\mlNotes
mlpSettings = %mlpDir%\settings.ini 
IfNotExist, %mlpSettings%
	{
	IniWrite, ERROR, %mlpDir%\settings.ini, Settings, Skin
	IniWrite, 1, %mlpDir%\settings.ini, Settings, NoteWin
	}	

mlpKolumner = %mlpDir%\kolumner.ini
FileAppend,,%mlpDir%\settings.ini
IniWrite, %version%, %mlpSettings%, Version, Version

IniRead, mainVersion, G:\NTM\NTM Digital Produktion\MedialinkPlus\dev\master.ini, Version, Version

; Ladda hem citat
gosub, citat


/*
-----------------------------------------------------------
	INSTALLATION
-----------------------------------------------------------
*/




/*
-----------------------------------------------------------
	UPPSTART & INITIALISERING
-----------------------------------------------------------
*/

; SPLASH
SplashImage, G:\NTM\NTM Digital Produktion\MedialinkPlus\dev\mlp2_42.jpg, B
Sleep 3000
SplashImage, Off

IfNotExist, %mlpKolumner%
	MsgBox, Du har inga inställningar för kolumner. Högerklicka på en order i Medialink och välj "MediaLink Plus -> Redigera Kolumner" för att ställa in.


; Läs inställningar
IniRead, noteWin, %mlpSettings%, Settings, NoteWin
IniRead, skin, %mlpSettings%, Settings, Skin


if (version < mainVersion){
	MsgBox, 68, Ny version!, Det finns en ny version av Medialink Plus. Vill du hämta den?
	IfMsgBox, Yes
		Run, G:\NTM\NTM Digital Produktion\MedialinkPlus
}

SetTimer, versionCheck, 5000
versionCheck:
	IniRead, mainVersion, G:\NTM\NTM Digital Produktion\MedialinkPlus\dev\master.ini, Version, Version
	if (version < mainVersion && traytip = 0){
		traytip = 1
		TrayTip, Ny version tillgänglig!, Det finns en ny version av Medialink Plus!`r`n Välj "Medialink Plus->Kontrollera uppdateringar" i menyn för att uppdatera.,10, 1
	}	

/*
-----------------------------------------------------------
	Extra-toolbars
-----------------------------------------------------------
*/

SetTimer, onTop, 100 ; Timer för extrafunktioner i Medialink-fönstret.

	onTop:
	if (noteWin = 1){
	IfWinActive, Atex MediaLink
	{
		mlActive = 1
		if (nyttcitat = 0)
		{
		WinGetTitle, mlTitle, Atex MediaLink
		mlNewTitle = %mlTitle%     ::     %citat% - %author%
		WinSetTitle, %mlNewTitle%
		nyttcitat = 1
		}

	} else {
		mlActive = 0
	}
	IfWinActive, noteWin
	{
		mlActive = 1
	}

	if (toolbar = 0 and mlActive = 1)
	{
		noteY =  0
		noteX =  0
		noteW =  100
		noteH =  100
		IniRead, noteX, %mlpDir%\notewin.ini, notewin, X
		IniRead, noteY, %mlpDir%\notewin.ini, notewin, Y
		IniRead, noteH, %mlpDir%\notewin.ini, notewin, H
		IniRead, noteW, %mlpDir%\notewin.ini, notewin, W
		if(!noteX || noteX = "ERROR")
			{
			noteX = 0
			}
		if(!noteY || noteY = "ERROR")
			{
			noteY = 0
			}
		if(!noteH || noteH = "ERROR")
			{
			noteH = 100
			}
		if(!noteW || noteW = "ERROR")
			{
			noteW = 100
			}
			
		Toolbar = 1
		WinGetPos, mlX, mlY,mlW,mlH,Atex MediaLink
		mlX7 := mlX + 790
		mlY7 := mlY + 54
		
		Gui, 7: Margin, 0, 0 
		Gui, 7: +ToolWindow -Caption ; no title, no taskbar icon 
		Gui, 7: Add, Picture, gcxMini ,%iconDir%\btn_cxmini.jpg
		Gui, 7: Show, x%mlX7% y%mlY7% NoActivate
		Gui, 7: +AlwaysOnTop
		
		Gui, 6: Margin, 0, 0
		Gui, 6: +ToolWindow -0x200000
		Gui, 6: Add, Edit, w220 h280 vinternaNoteringarEdit
		Gui, 6:+Resize -MaximizeBox
		Gui, 6: Show, x%noteX% y%noteY% h%noteH% w%noteW% NoActivate, noteWin
		Gui, 6: +AlwaysOnTop
		
		6GuiSize: 
			GuiControl Move, internaNoteringarEdit, % "H" . (A_GuiHeight) . " W" . (A_GuiWidth)
			If A_EventInfo = 1  ; The window has been minimized.  No action needed.
			Return
			if (A_EventInfo = 0)
			{
				
				WinGetPos,x,y,w,h, noteWin
				if (h != 0)
				{
				h := h-34
				}
				if (w != 0)
				{
				w := w-16
				}
				IniWrite, %x%, %mlpDir%\notewin.ini, notewin, X
				IniWrite, %y%, %mlpDir%\notewin.ini, notewin, Y
				IniWrite, %h%, %mlpDir%\notewin.ini, notewin, H
				IniWrite, %w%, %mlpDir%\notewin.ini, notewin, W
			}
		Return
		
		
	}
	if (toolbar = 0 and mlActive = 0)
	{
		toolbar = 0
		Gui, 7: Destroy
		Gui, 6: Destroy
	}
	if (toolbar = 1 och mlActive = 1)
	{
		return
	}
	if (toolbar = 1 and mlActive = 0)
	{
		toolbar = 0
		Gui, 7: Destroy
		Gui, 6: Destroy
	}
	return
}
return

/*
-----------------------------------------------------------
	MENY
-----------------------------------------------------------
*/
~LButton::
	IfWinActive, Atex MediaLink
	{
		MouseGetPos, , , id, control
		IfInString, control, SysListView ; kontrollerar att man klickat i en SysListView
		{
			SLV = 1
		} else {
			SLV = 0
		}

		If (SLV =1)
		{ 
			temp := Clipboard
			Send, ^c
			mlOrdernummer := Clipboard
			Clipboard := temp
			sleep, 400
			ControlGetText, internaNoteringar, Edit3, Atex MediaLink
			; Kontrollera om interna noteringar uppdaterats
			;~ IfExist, %notesDir%\%mlOrdernummer%.txt
				;~ FileRead, mlNote, %notesDir%\%mlOrdernummer%.txt
			;~ IfExist, %notesDir%\%mlOrdernummer%.txt
				;~ ControlSetText, Edit3, %internaNoteringar%`r`r%mlNote%, Atex MediaLink
			;~ IfExist, %notesDir%\%mlOrdernummer%.txt
				;~ ControlSetText, Edit1, %internaNoteringar%`r`r%mlNote%, noteWin 
			;~ IfNotExist,  %notesDir%\%mlOrdernummer%.txt
				ControlSetText, Edit1, %internaNoteringar%, noteWin 
			
			
		}
	}	
return

~RButton::
	MouseGetPos, , , id, control
	IfInString, control, SysListView ; kontrollerar att man klickat i en SysListView
	{
		SLV = 1
	} else {
		SLV = 0
	}

	If (WinActive("ahk_class wxWindowClassNR") and SLV = 1)
	{
		Click, right
		Send, {Esc}
		if (menuOn = 1)
		{
			menu, Menu, DeleteAll
		}
		; Huvudmenyn
		; Printcheck
		gosub, getList
		printCheck(mlOrdernummer) ; Kolla om markerad order har en printannons klar på mtrl -01
		menu, Menu, add, &Hitta Print-PDF, printCheck
		menu, Menu, Icon, &Hitta Print-PDF, %iconDir%\inteprint.ico
		menu, Menu, disable, &Hitta Print-PDF
		if (printFinns = 1) ; Om print finns
		{
			menu, Menu, Icon, &Hitta Print-PDF, %iconDir%\print.ico
			menu, Menu, enable, &Hitta Print-PDF
		}

		; cxense-menyn
		menu, cxense, add, Boka kampanj, cxStart
		menu, cxense, add, Öppna i cxense, oppnaCXfork
		menu, cxense, add, Rapport, oppnaRapportFork
		menu, cxense, Icon, Boka kampanj, %iconDir%\boka.ico
		menu, cxense, Icon, Öppna i cxense, %iconDir%\oppna.ico
		menu, cxense, Icon, Rapport, %iconDir%\cxrapport.ico

		menu, Menu, add, Kopiera kundnamn och ordernummer, kundOrder
		menu, Menu, add, Sök på ordernummer, sokOrder
		menu, Menu, add, Tilldela och Bearbeta, statusTilldelaBearbetas
		menu, Menu, add, Cxense, :cxense
		menu, Menu, add, Starta annons i Photoshop, startaAnnonsPS
		menu, Menu, add, Starta annons i Flash, startaAnnonsFlash
		menu, Menu, add, Öppna kundmapp, oppnaKundmapp
		menu, Menu, add


		; Ikoner för ovanstående
		menu, Menu, Icon, Kopiera kundnamn och ordernummer, %iconDir%\kopiera.ico
		menu, Menu, Icon, Sök på ordernummer, %iconDir%\sok.ico
		menu, Menu, Icon, Tilldela och Bearbeta, %iconDir%\tilldelabearbeta.ico
		menu, Menu, Icon, Starta annons i Photoshop, %iconDir%\photoshop.ico
		menu, Menu, Icon, Starta annons i Flash, %iconDir%\flash.ico
		menu, Menu, Icon, Öppna kundmapp, %iconDir%\kundmapp.ico
		menu, Menu, Icon, Cxense, %iconDir%\cx.ico
		
		; Status-menyn
		menu, status, add, Klar, statusKlar
		menu, status, add, Obekräftad, statusObekraftad
		menu, status, add, Bearbetas, statusBearbetas
		menu, status, add, Korrektur skickat, statusKorrSkickat
		menu, status, add, Korrektur klart, statusKorrekturKlart
		menu, status, add, Vilande, statusVilande
		menu, status, add, Manus på Mail, statusManusMail
		menu, status, add, Undersöks, statusUndersoks
		menu, status, add, Arkiverad, statusArkiverad
		menu, status, add, Ny, statusNy
		menu, status, add, Repetition, statusRep
		menu, status, add, Lev. Färdig, statusLevFardig
		menu, status, add, Bokad, statusBokad
		menu, status, add, Ej komplett manus, statusEjkomplett
		menu, Menu, add, Status, :status

		; Status-menyn - Ikoner
		menu, status, Icon, Klar, %iconDir%\klar.ico
		menu, status, Icon, Obekräftad, %iconDir%\klar.ico
		menu, status, Icon, Bearbetas, %iconDir%\bearbetas.ico
		menu, status, Icon, Korrektur skickat, %iconDir%\korr.ico
		menu, status, Icon, Korrektur klart, %iconDir%\korr.ico
		menu, status, Icon, Vilande, %iconDir%\vilande.ico
		menu, status, Icon, Manus på Mail, %iconDir%\vilande.ico
		menu, status, Icon, Undersöks, %iconDir%\undersoks.ico
		menu, status, Icon, Arkiverad, %iconDir%\arkiverad.ico
		menu, status, Icon, Ny, %iconDir%\ny.ico
		menu, status, Icon, Repetition, %iconDir%\rep.ico
		menu, status, Icon, Lev. Färdig, %iconDir%\fardig.ico
		menu, status, Icon, Bokad, %iconDir%\ny.ico
		menu, status, Icon, Ej komplett manus, %iconDir%\ny.ico
		menu, Menu, Icon, Status, %iconDir%\status.ico

		; Tilldela-menyn
		menu, tilldela, add, Mig, tilldelaMig
		menu, tilldela, add, Annan, tilldelaAnnan
		menu, tilldela, add, Ingen, tilldelaIngen
		menu, Menu, add, Tilldela..., :tilldela
		menu, Menu, Icon, Tilldela..., %iconDir%\tilldela.ico
		menu, tilldela, Icon, Mig, %iconDir%\mig.ico
		menu, tilldela, Icon, Annan, %iconDir%\annan.ico
		menu, Menu, add

		; Undre menyn
		menu, Menu, add, Maila säljare, saljarMail
		menu, Menu, add, Maila korrektur, korrMail
		menu, Menu, add, Rapportera felaktig order, rapporteraFel
		menu, Menu, add

		; Undre meyn - Ikoner
		menu, Menu, Icon, Maila säljare, %iconDir%\mail.ico
		menu, Menu, Icon, Maila korrektur, %iconDir%\mail.ico
		menu, Menu, Icon, Rapportera felaktig order, %iconDir%\rapportera.ico
		
		; Traffic-menyn
		menu, traffic, add, Leta printannonser, kontrolleraPrint
		menu, traffic, add, Kopiera kampanjer, copyCampaigns
		menu, traffic, add, Räkna exponeringar, raknaExponeringar
		menu, traffic, add, Räkna antal markade annonser, raknaValda
		menu, traffic, add, Uppdatera lager-XML, lager
		menu, Menu, add, Traffic, :traffic
		menu, Menu, Icon, Traffic, %iconDir%\traffic.ico
		
		; MediaLink Plus-menyn
		menu, mlp, add, Redigera kolumner, kolumner
		menu, mlp, add, Kontrollera uppdateringar, updateCheck
		menu, mlp, add, Uppdatera citat, cycleCitat
		menu, mlp, add, Inställningar, settings
		menu, Menu, add, MediaLink Plus, :mlp
		menu, Menu, Icon, MediaLink Plus, %iconDir%\plus.ico
		menu, mlp, Icon, Redigera kolumner, %iconDir%\columns.ico
		menu, mlp, Icon, Kontrollera uppdateringar, %iconDir%\uppdatera.ico
		menu, mlp, Icon, Inställningar, %iconDir%\settings.ico

		menu, Menu, Color, FFFFFF 
		menu, Menu, show

		menuOn = 1
	}
return



/*
-----------------------------------------------------------
	FUNKTIONER
-----------------------------------------------------------
*/

printCheck(x)
{
	global
	StringTrimRight, OrderNummerUtanMnr, x, 3 ; Tar bort tre sista tecken i ordernumret, dvs materialnummer inkl bindestreck
	StringTrimLeft, SistaTvaSiffrorna, OrderNummerUtanMnr, 8 ; Plockar ut sista två siffrorna ur ordernumret
	StringTrimLeft, OrderNummerUtanNollor, OrderNummerUtanMnr, 3 ; Tar bort inledande nollor i ordernumret
	sokVag = \\nt.se\Adbase\Annonser\Ad\%SistaTvaSiffrorna%\10%OrderNummerUtanNollor%-01.pdf
	printFinns = 0
	ifExist, %sokVag%
		printFinns = 1
	return
}

rensaTecken(ByRef x) ; Rensar ur valda tecken ur en variabel
{
		StringReplace, x, x, &&, &, All
		StringReplace, x, x, /,%A_SPACE%, All
		StringReplace, x, x, \,%A_SPACE%, All
		StringReplace, x, x, :,%A_SPACE%, All
}
stripDash(ByRef x)
{
	StringReplace, x, x,-,,All
}

status(x) ; Sätter status enligt x
{
	Send, !s
	WinWaitActive, Change Status
	Send, {Tab}
	Control, ChooseString, %x%, ComboBox1
	Send, {Tab}{Enter}
}

assign(x) ; tilldelar till x 
{
	Send, !a
	WinWaitActive, Ändra tilldelad
	Send, {Tab}
	Control, ChooseString, %x%, ComboBox1
	Send, {Tab}{Enter}
}

mailTo(saljare, subject, ordernummer, kundnamn)
{
	SetKeyDelay, 0
	WinActivate, Microsoft Outlook
	Run, C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE /c IPM.Note
	WinWaitActive, Namnlös - Meddelande
	Send, %saljare%
	Send, !m
	Send, %subject% %kundnamn% (%ordernummer%){Tab}
}

getFormat(x)
{
	global
	;Skyltar
	if (x = "Eurosize")
		{
		format = 1080x1920
		file = 1080 x 1920
		}
	if (x = "Söderköpingsvägen")
		{
		format = 1024x384	
		file = 1024 x 384
		}
	if (x = "Hamnbron")
		{
		format = 1024x384
		file = 1024 x 384
		}
	if (x = "Sjötull")
		{
		format = 512x128
		file = 512 x 128
		}
	if (x = "Skylt HD")
		{
		format = 1920x1080
		file = 1920 x 1080
		}
	if (x = "Ståthögaleden")
		{
		format = 768x384
		file = 768 x 384
		}
	if (x = "Östcentrum Visby")
		{
		format = 1920x720
		file = 1920 x 720
		}

	; Moduler
	if (x = "Artikel 120")
	{
		format = MOD
		file = 468 x 120
	}
	if (x = "Mittbanner 1")
	{
		format = MOD
		file = 468 x 240
	}
	if (x = "Modul 240")
	{
		format = MOD
		file = 468 x 240
	}
	if (x = "Modul MK")
	{
		format = MKMOD
	}
	
	; Mobiler
	if (x = "Mobil Bottom Panorama")
	{
		format = MOB
		file = 320 x 80
	}
	if (x = "Mobil Bottom Panorama XL")
	{
		format = MOB
		file = 320 x 160
	}
	if (x = "Mobil Bottom Takeover")
	{
		format = MOB
		file = 320 x 320
	}
	if (x = "Mobil Top Panorama")
	{
		format = MOB
		file = 320 x 80
	}
	if (x = "Mobil Top Panorama XL")
	{
		format = MOB
		file = 320 x 160
	}
	if (x = "Mobil Top Takeover")
	{
		format = MOB
		file = 320 x 320
	}
	if (x = "Mobil Stor")
	{
		format = MOB
		file = 320 x 320
	}
	if (x = "Mobil Mellan")
	{
		format = MOB
		file = 320 x 160
	}
	if (x = "Mobil Liten")
	{
		format = MOB
		file = 320 x 80
	}
	if (x = "Mobil Swipe")
	{
		format = MOB
		file = 320 x 320
	}

	; Outsider
	if (x = "Outsider 600")
	{
		format = OUT
		file = 250 x 600
	}
	if (x = "Outsider 800")
	{
		format = OUT
		file = 250 x 800
	}
	if (x = "Skyskrapa")
	{
		format = OUT
		file = 250 x 600
	}

	; Panorama
	if (x = "Panorama")
	{
		format = PAN
		file = 980 x 240
	}
	if (x = "Panorama 1 120")
	{
		format = PAN
		file = 980 x 120
	}
	if (x = "Panorama 1 240")
	{
		format = PAN	
		file = 980 x 240
	}
	
	if (x = "Panorama 480")
	{
		format = PAN	
		file = 980 x 480
	}
	
	if (x = "Panorama 120")
	{
		format = PAN
		file = 980 x 120
	}
	if (x = "Panorama 2 120")
	{
		format = PAN
		file = 980 x 120
	}
	if (x = "Panorama 2 240")
	{
		format = PAN
		file = 980 x 240
	}
	if (x = "Panorama 2 480")
	{
		format = PAN
		file = 980 x 480
	}
	if (x = "Panorama 240")
	{
		format = PAN
		file = 980 x 240
	}
	if (x = "Panorama 360")
	{
		format = PAN
		file = 980 x 360
	}

	if (x = "Portal 980")
	{
		format = 980
		file = 980 x 360
	}



	; Widescreen
	if (x = "Widescreen 240")
	{
		format = WID
		file = 250 x 240
	}

	; Portaler
	if (x = "Kvadrat")
	{
		format = 180
		file = 180 x 180
	}
	if (x = "Portal 180")
	{
		format = 180
		file = 180 x 180
	}
	if (x = "Stortavla")
	{
		format = 380
		file = 380 x 280
	}

	if (x = "Portal 380")
	{
		format = 380
		file = 380 x 280
	}

	if (x = "Portal 580")
	{
		format = 580
		file = 580 x 280
	}
	if (x = "Textannons")
	{
		format = TXT
	}

	; Reach
	if (x = "Reach 250")
	{
		format = REACH250
		file = 250 x 360
	}
	if (x = "Reach 468")
	{
		format = REACH468
		file = 468 x 240
	}

	; Väder
	if (x = "Väderspons")
	{
		format = VADER
	}
	
	if (x = "Julkalender")
	{
		format = MOB
	}

	if (x = "Helsida")
	{
		format = HEL
	}
}

cxProduct(format, type)
{
	global
		;---- Riktade
	RiktadMOD = 00000001609df500
	RiktadOUT = 00000001609df509
	RiktadPAN = 00000001609df502
	RiktadWID = 00000001609df506
	Riktad380 = 0000000160ec7963
	Riktad180 = 0000000160eb55e5
	RiktadTXT = 0000000160f4b7d6

	;---- ROS
	RosMOD = 0000000160209522
	RosOUT = 0000000160209517
	RosPAN = 00000001601d3817
	RosWID = 0000000160209515
	Ros580 = 00000001610811d9
	Ros380 = 0000000160ec7931
	Ros180 = 0000000160da016b
	RosTXT = 0000000160f4b805

	;---- PLUGG
	PluggMOD = 00000001613e2960
	PluggOUT = 00000001613e2966
	PluggPAN = 00000001613e296c
	PluggWID = 00000001613e2c51
	Plugg380 = 00000001613e2944
	Plugg180 = 00000001613e291b
	PluggTXT = 0000000160f4b848
	PluggMOB = 000000016116a6d6

	;---- REACH
	Reach = 000000015f460c65

	;---- MOBIL
	Mobil = 00000001608fa111

	;---- RETARGET
	RetargetMob = 0000000160fb5848
	RetargetPan = 0000000160fb4fd9
	
	;---- CPC
	CPC380 = 000000016107992d
	CPC180 = 00000001610799f5
	CPCTXT = 0000000161079a11
	PluggCPC380 = 00000001610798f2
	PluggCPC180 = 0000000161079a00
	PluggCPCTXT = 0000000161079a17
	CPC980 = 000000016108d7ba
	
	;---- Modul MK
	MKMOD = 0000000161107ccf

	;---- Helsida
	HELid = 0000000160a381fe
	
	if (format = "MOD" && type = "Run On Site")
	{
		productID = %RosMOD%
		template = 000000016020bedd
	}
	
	if (format = "OUT" && type = "Run On Site")
	{
		productID = %RosOUT%
		template = 000000016020ba25
	}
	
	if (format = "WID" && type = "Run On Site")
	{
		productID = %RosWID%
		template = 000000016020bf37
	}
	
	if (format = "PAN" && type = "Run On Site")
	{
		productID = %RosPAN%
		template = 000000016020ba73
	}

	if (format = "380" && type = "Run On Site")
	{
		productID = %Ros380%
		template = 0000000161079935
	}
	
	if (format = "180" && type = "Run On Site")
	{
		productID = %Ros180%
		template = 00000001610799fe
	}

	if (format = "MOD" && type = "Riktad")
	{
		productID = %RiktadMOD%
		template = 00000001609e56da
	}
	
	if (format = "OUT" && type = "Riktad")
	{
		productID = %RiktadOUT%
		template = 00000001609e5542
	}

	if (format = "WID" && type = "Riktad")
	{
		productID = %RiktadWID%
		template = 00000001609ea38e
	}
	
	if (format = "PAN" && type = "Riktad")
	{
		productID = %RiktadPAN%
		template = 00000001609ea651
	}
	
	if (format = "380" && type = "Riktad")
	{
		productID = %Riktad380%
		template = 
	}

	if (format = "980")
	{
		productID = %CPC980%
		template = 000000016108d84c
	}


	if (format = "180" && type = "Riktad")
	{
		productID = %Riktad180%
		template = 
	}
	
	if (format = "MOD" && type = "Plugg")
	{
		productID = %PluggMOD%
		template = 00000001613e2963
	}

	if (format = "MOB" )
	{
		productID = %Mobil%
		template = 0000000160fb5086
	}

	if (format = "MOB" && type = "Plugg")
	{
		productID = %PluggMOB%
		template = 00000001613f727a
	}

	if (format = "OUT" && type = "Plugg")
	{
		productID = %PluggOUT%
		template = 00000001613e2969
	}
	
	if (format = "WID" && type = "Plugg")
	{
		productID = %PluggWID%
		template = 00000001613e2c54
	}

	if (format = "PAN" && type = "Plugg")
	{
		productID = %PluggPAN%
		template = 00000001613e297d
	}

	if (format = "380" && type = "Plugg")
	{
		productID = %Plugg380%
		template = 00000001613e2948
	}

	if (format = "180" && type = "Plugg")
	{
		productID = %Plugg180%
		template = 00000001613e291e
	}

	if (format = "380" && type = "CPC")
	{
		productID = %CPC380%
		template = 
	}

	if (format = "180" && type = "CPC")
	{
		productID = %CPC180%
		template = 
	}

	if (format = "PAN" && type = "CPC")
	{
		productID = %CPC180%
		template = 000000016108d7ba
	}
	
	if (format = "380" && type = "CPC Plugg")
	{
		productID = %PluggCPC380%
		template = 000000016108d84c
	}

	if (format = "180" && type = "CPC Plugg")
	{
		productID = %PluggCPC180%
		template = 0000000161079a0f
	}


	if (format = "REACH468")
	{
		productID = %Reach%
		template = 000000016017a451
	}

	if (format = "TXT" && type = "Run On Site")
	{
		productID = %RosTXT%
		template = 0000000161079a15
	}

	if (format = "TXT" && type = "CPC")
	{
		productID = %CPCTXT%
		template = 0000000161079a15
	}

	if (format = "TXT" && type = "CPC Plugg")
	{
		productID = %PluggCPCTXT%
		template = 
	}

	if (format = "MOB" && type = "Retarget")
	{
		productID = %RetargetMob%
		template = 
	}

	if (format = "PAN" && type = "Retarget")
	{
		productID = %RetargetPan%
		template = 
	}

	if (format = "580")
	{
		productID = %Ros580%
		template = 00000001610811d9
	}

	if (format = "WID" and mlTidning = "AF")
	{
		productID = %Ros380%
		template = 
	}
	
	if (format = "MKMOD")
	{
		productID = %MKMOD%
		template = 
	}

	if (format = "HEL")
	{
		productID = %HELid%
		template = 0000000160a38246
	}
	
}

forstaBokstav(x)
{
	StringLen, tecken, x
	tecken := tecken - 1
	StringTrimRight, forstaBokstav, x, %tecken%
	return forstaBokstav
}

refreshFile(content, file)
{
	FileDelete, %file%
	FileEncoding, UTF-8-RAW
	FileAppend, %content%, %file%
	FileEncoding
}

/*
-----------------------------------------------------------
	SUBRUTINER
-----------------------------------------------------------
*/

getAnvnamn:
	WinGetTitle, Windowtext, Atex MediaLink
	StringSplit, WindowSplit, Windowtext, =
	StringSplit, WindowSplit, WindowSplit2, %A_Space%
	Anvandare =  %WindowSplit1%
	StringTrimRight, AnvKort, Anvandare, 2 ; Sätter AnvKort till användarens förnamn

return

getList: ; Hämtar information från valt objekt i listvyn
	gosub, getAnvnamn
	ControlGet, listCount, List, Count Selected, %control%, Atex MediaLink
	ControlGet, getList, List, Selected, %control%, Atex MediaLink
	if (weblinkget = 1)
		ControlGet, getList, List, , %control%, Atex MediaLink
	StringSplit, getListRow, getList, `n
	listRow = %getListRow1%
	Stringsplit, kolumn, listRow, `t

	;Läs kolumn-info från användarens kolumner.ini
	IniRead, iniStart, %mlpDir%\kolumner.ini, kolumner, Start
	IniRead, iniStopp, %mlpDir%\kolumner.ini, kolumner, Stopp
	IniRead, iniExponeringar, %mlpDir%\kolumner.ini, kolumner, Exponeringar
	IniRead, iniKundnr, %mlpDir%\kolumner.ini, kolumner, Kundnr
	IniRead, iniKundnamn, %mlpDir%\kolumner.ini, kolumner, Kundnamn
	IniRead, iniSaljare, %mlpDir%\kolumner.ini, kolumner, Saljare
	IniRead, iniProdukt, %mlpDir%\kolumner.ini, kolumner, Produkt
	IniRead, iniEnhet, %mlpDir%\kolumner.ini, kolumner, Internetenhet
	IniRead, iniStatus, %mlpDir%\kolumner.ini, kolumner, Status
	IniRead, iniTilldelad, %mlpDir%\kolumner.ini, kolumner, Tilldelad


	mlStartdatum := kolumn%iniStart%
	mlStoppdatum := kolumn%iniStopp%
	mlExponeringar := kolumn%iniExponeringar%
	mlKundnr := kolumn%iniKundnr%
	mlKundnamn := kolumn%iniKundnamn%
	mlSaljare := kolumn%iniSaljare%
	mlProdukt := kolumn%iniProdukt%
	mlStatus := kolumn%iniStatus%
	mlTilldelad := kolumn%iniTilldelad%
	
	StringSplit, prodArray, mlProdukt , %A_Space%
	mlTidning = %prodArray1%
	mlSite = %prodArray2%

	if (mlSite = "gotland.net")
	{
		mlTidning = GN
	}
	if (mlSite = "nt.se")
	{
		mlTidning = NTFB
	}

	mlEnhet := kolumn%iniEnhet%
	mlOrdernummer = %kolumn1%
	weblinkget = 0
	
return

printCheck:
	printCheck(mlOrdernummer) ; printFinns = 1 eller 0
	if (printFinns = 1)
	{
		run, %sokVag%
	}
return

kundOrder:
	kundOrder = %mlKundnamn% (%mlOrdernummer%)
	clipboard := kundOrder
	ToolTip, Kopierat: %KundOrder%
	SetTimer, RemoveToolTip, 1500
return

RemoveToolTip:
	SetTimer, RemoveToolTip, Off
	ToolTip
return

Kolumner: ; Redigerar ordningen på användarens kolumner och sparar den i kolumner.ini
		kolumnLista =
	(
1: %kolumn1%
2: %kolumn2%
3: %kolumn3%
4: %kolumn4%
5: %kolumn5%
6: %kolumn6%
7: %kolumn7%
8: %kolumn8%
9: %kolumn9%
10: %kolumn10%
11: %kolumn11%
12: %kolumn12%
13: %kolumn13%
14: %kolumn14%
15: %kolumn15%
16: %kolumn16%
17: %kolumn17%
18: %kolumn18%
19: %kolumn19%
20: %kolumn20%
	)

; öppnar dialogruta
;Gui, Add, GroupBox, x2 y0 w230 h300 , Kolumner
Gui, Add, Text, x12 y20 w210 h270 , %kolumnLista%
Gui, Add, Text, x12 y310 w90 h20 , Startdatum:
Gui, Add, Text, x12 y340 w90 h20 , Stoppdatum:
Gui, Add, Text, x12 y370 w90 h20 , Exponeringar:
Gui, Add, Text, x12 y400 w90 h20 , Kundnr:
Gui, Add, Text, x12 y430 w90 h20 , Kundnamn:
Gui, Add, Text, x12 y460 w90 h20 , Faktisk Säljare:
Gui, Add, Text, x12 y490 w90 h20 , Produkt:
Gui, Add, Text, x12 y520 w90 h20 , Internetenhet:
Gui, Add, Text, x12 y550 w90 h20 , Status:
Gui, Add, Text, x12 y580 w90 h20 , Tilldelad:
Gui, Add, DropDownList, x132 y310 w100 h20 R20 Choose%iniStart% vddStart, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y340 w100 h20 R20 Choose%iniStopp% vddStopp, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20 
Gui, Add, DropDownList, x132 y370 w100 h20 R20 Choose%iniExponeringar% vddExp, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20 
Gui, Add, DropDownList, x132 y400 w100 h20 R20 Choose%iniKundnr% vddKnr, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20 
Gui, Add, DropDownList, x132 y430 w100 h20 R20 Choose%iniKundnamn% vddKnamn, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y460 w100 h20 R20 Choose%iniSaljare% vddSaljare, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y490 w100 h20 R20 Choose%iniProdukt% vddProdukt, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y520 w100 h20 R20 Choose%iniEnhet% vddInternetenhet, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y550 w100 h20 R20 Choose%iniStatus% vddStatus, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, DropDownList, x132 y580 w100 h20 R20 Choose%iniTilldelad% vddTilldelad, 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20
Gui, Add, Button, x12 y610 w100 h30 gkolumnSubmit, Spara
Gui, Add, Button, x132 y610 w100 h30 gAvbryt, Avbryt
Gui, Color, FFFFFF
Gui, +ToolWindow +Caption
Gui, Show, x843 y311 h650 w244, Kolumn-konfigurering

Return


Avbryt:
Gui, Destroy
return

kolumnSubmit: ; Sparar information från dialogruta skapad i sub Kolumner
	Gui, Submit
	Gui, Destroy
	IniDelete, %mlpDir%\kolumner.ini, Kolumner
	IniWrite, %ddStart%, %mlpDir%\kolumner.ini, Kolumner, Start
	IniWrite, %ddStopp%, %mlpDir%\kolumner.ini, Kolumner, Stopp
	IniWrite, %ddExp%, %mlpDir%\kolumner.ini, Kolumner, Exponeringar
	IniWrite, %ddKnr%, %mlpDir%\kolumner.ini, Kolumner, Kundnr
	IniWrite, %ddKnamn%, %mlpDir%\kolumner.ini, Kolumner, Kundnamn
	IniWrite, %ddSaljare%, %mlpDir%\kolumner.ini, Kolumner, Saljare
	IniWrite, %ddProdukt%, %mlpDir%\kolumner.ini, Kolumner, Produkt
	IniWrite, %ddInternetenhet%, %mlpDir%\kolumner.ini, Kolumner, Internetenhet
	IniWrite, %ddStatus%, %mlpDir%\kolumner.ini, Kolumner, Status
	IniWrite, %ddTilldelad%, %mlpDir%\kolumner.ini, Kolumner, Tilldelad
Return

sokOrder:
	StringTrimRight, OrderNummerUtanMnr, mlOrdernummer, 3 ; Tar bort tre sista tecken i ordernumret, dvs materialnummer inkl bindestreck
	ControlSetText, Edit1, %OrderNummerUtanMnr%, Atex MediaLink ; Sätt ovanstående i sökfältet
	ControlFocus, Edit1, Atex MediaLink	; Sätter fokus på sökfältet
	Send, {Enter} ; Trycker enter för att starta sök.
return

saljarMail:
	assign(AnvKort)
	status("Undersöks")
	subject = Fråga: ; Text att sätta i ämnesraden innan kundnamn (ordernummer)
	mailTo(mlSaljare, subject, mlOrdernummer, mlKundnamn)
Return

korrMail:
	assign(AnvKort)
	status("Korrektur skickat")
	order = 1
	bokning = 0
	gosub, cxGetAdId
	if (order = 1)
	{
	subject = Korrektur: ; Text att sätta i ämnesraden innan kundnamn (ordernummer)
	mailTo(mlSaljare, subject, mlOrdernummer, mlKundnamn)
	text =
	(


----

{CTRL down}f{CTRL up}Länk för rapport:{CTRL down}f{CTRL up}
http://rapport.ntm-digital.se/advertiser/%kundID%/campaign/%campaignID%/

Spara denna länk{SHIFT down}1{SHIFT up}
På denna länk finns information om hur denna order är inbokad. Kontrollera så att start/stoppdatum stämmer och att antal exponeringar är korrekt. Om vi inte får något svar startar denna annons på startdatumet. Denna länk fungerar även som statusrapport för denna kampanj. Här kan du se exakt hur många exponeringar/klick som levererats hittills i kampanjen. Om du någon gång under kampanjens gång upplever att något inte står rätt till, kontakta Traffic.

----

{CTRL down}{Home}{CTRL up}
	)
	Send, %text%
}
if (order = 0)
{
	subject = Korrektur: ; Text att sätta i ämnesraden innan kundnamn (ordernummer)
	mailTo(mlSaljare, subject, mlOrdernummer, mlKundnamn)
}
Return

startaAnnonsPS:
	status("bearbetas") ; sätter status bearbetas
	assign(AnvKort) ; tilldelar till användaren
	getFormat(mlEnhet) ; Hämtar formatet utifrån internetenhet
	stripDash(mlStartdatum) ; Tar bort - ur startdatum
	rensaTecken(mlKundnamn) 
	StringTrimLeft, mlStartdatum, mlStartdatum, 2 ; tar bort första två tecknen ur datumet
	forstaBokstav := forstaBokstav(mlKundnamn)

	adDir = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstav%\%mlKundnamn%\%mlStartdatum%
	if FileExist(adDir)
	{
		FileCopy, %templateDir%\%file%.psd, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.psd
		run, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.psd
	} else {
		FileCreateDir, %adDir%
		FileCopy, %templateDir%\%file%.psd, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.psd
		run, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.psd
	}
return

startaAnnonsFlash:
	status("bearbetas") ; sätter status bearbetas
	assign(AnvKort) ; tilldelar till användaren
	getFormat(mlEnhet) ; Hämtar formatet utifrån internetenhet
	stripDash(mlStartdatum) ; Tar bort - ur startdatum
	rensaTecken(mlKundnamn) 
	StringTrimLeft, mlStartdatum, mlStartdatum, 2 ; tar bort första två tecknen ur datumet
	forstaBokstav := forstaBokstav(mlKundnamn)

	adDir = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstav%\%mlKundnamn%\%mlStartdatum%
	if FileExist(adDir)
	{
		FileCopy, %templateDir%\%file%.fla, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.fla
		run, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.fla
	} else {
		FileCreateDir, %adDir%
		FileCopy, %templateDir%\%file%.fla, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.fla
		run, %adDir%\%mlTidning%%format%-%mlKundnamn%-%mlStartdatum%.fla
	}
return

oppnaKundmapp:
	getFormat(mlEnhet) ; Hämtar formatet utifrån internetenhet
	stripDash(mlStartdatum) ; Tar bort - ur startdatum
	rensaTecken(mlKundnamn) 
	StringTrimLeft, mlStartdatum, mlStartdatum, 2 ; tar bort första två tecknen ur datumet
	forstaBokstav := forstaBokstav(mlKundnamn)

	adDir = G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstav%\%mlKundnamn%\
	ifExist, %adDir%
	{
		run, G:\NTM\NTM Digital Produktion\Webbannonser\0-Arkiv\%A_YYYY%\%forstaBokstav%\%mlKundnamn%\
	} else {
		msgbox, Ingen mapp hittades på denna sökväg:`r`n%adDir%
	}
return

kontrolleraPrint:
StatusBarGetText, xAds, 3, Atex MediaLink
StringSplit, xAds, xAds, %A_Space%
xAds = %xAds1%
msgbox, Vill du gå igenom %xAds% annonser?
adCount = 0
printAds = 0
while (adCount < xAds)
{
	adCount++
	Sleep, 50
	Send, ^c
	Sleep, 50
	printFinns = 0
	printCheck(Clipboard) ; printFinns = 1 eller 0
		if (printFinns = 1)
		{
			printAds++
			status("Vilande")
		}
	Sleep, 200
	Send, {Down}
}
msgbox, Klar!`n%printAds% annonser av %xAds% hade printmaterial tillgängligt.
return

updateCheck:
IniRead, mainVersion, G:\NTM\NTM Digital Produktion\MedialinkPlus\dev\master.ini, Version, Version
if (version < mainVersion){
	MsgBox, 68, Ny version!, Det finns en ny version av Medialink Plus. Vill du hämta den?
	IfMsgBox, Yes
		{
			IfNotExist, %A_MyDocuments%\MedialinkPlus\Temp
				FileCreateDir, %A_MyDocuments%\MedialinkPlus\Temp
			FileDelete, %A_MyDocuments%\MedialinkPlus\Temp\install.exe
			UrlDownloadToFile, http://dnns.se/mlp/mlp/latest/install.exe, %A_MyDocuments%\MedialinkPlus\Temp\install.exe

			If (ErrorLevel != 1) ; Om filen laddades ner
			{
				Run, %A_MyDocuments%\MedialinkPlus\Temp\install.exe
			}
		}
}
return
/*
-----------------------------------------------------------
	CX-Bokning
-----------------------------------------------------------
*/

cxStart:
	if (listCount > 1)
	{
		goto, bokaFlera
	} else {
		goto, cxBokning
	}
return

cxBokning:
	Send, ^c
	bokning = 1
	kundID = 
	goto, cxKundCheck ; returnerar kundID
return

cxCPM:
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

cxSplashOn:
	SplashImage, G:\NTM\NTM Digital Produktion\övrigt\MedialinkPlus\dev\cx_loading.jpg, B
return

cxSplashOff:
	SplashImage, Off
return

cxKundCheck:
	kundCheck = -%A_Space%%mlKundnr%%A_Space%-
	xmlSection = 
	xmlPart = 0
	gosub, cxSplashOn
	; --------------------------- HTTP-Request ---------------------------
	URL = https://%cxUser%@cxad.cxense.com/api/secure/folder/advertising
	DATA := ""
	HEAD := ""
	OPTS = SaveAs: %cxDir%\kundCheck.xml`nCharset: utf-8
	HTTPRequest( URL, DATA, HEAD, OPTS )
	; --------------------------------------------------------------------

	FileRead, kundXML, %cxDir%\kundCheck.xml
	StringReplace, kundXML, kundXML, <cx:childFolder>, +, A

	; Loopa igenom XML-filen för att se om kund finns
	Loop, Parse, kundXML, +
	{
		if InStr(A_LoopField, kundCheck)
		{
			xmlPart = %A_LoopField%
			break
		}
	}
	if (xmlPart != 0) ; Kund hittad
	{
		StringSplit, xmlSplit, xmlPart, >
		StringSplit, xmlSplit, xmlSplit4, <
		kundID = %xmlSplit1% ; kundID innehåller kundens ID
		gosub, cxSplashoff
		if (bokning != 0)
		{
		goto, cxKampanjbokning
		}
	}
	
	if (xmlPart = 0 && bokning = 0) ; Kund inte hittad, ingen inbokning
	{
		gosub, cxSplashOff
		MsgBox, Kunde inte hitta kund, avbryter.
		order = 0
		return
	}
	if (xmlPart = 0) ; Kund inte hittad
	{
		gosub, cxSplashoff
		goto, cxKundSaknas
		return
	}
	gosub, cxSplashOff
return

cxKundSaknas:
	MsgBox, 4, Kund saknas, Kund fanns inte. Skapa kund "%mlTidning% - %mlKundnr% - %mlKundnamn%"?
	IfMsgBox, Yes 
	{
		gosub, cxSplashOn
		StringReplace, mlKundnamn, mlKundnamn,&,,A
		; --------------------------- HTTP-Request ---------------------------
		URL = https://%cxUser%@cxad.cxense.com/api/secure/folder/advertising
		DATA =
		(
		<?xml version="1.0" encoding="UTF-8"?>
		<cx:folder xmlns:cx="http://cxense.com/cxad/api/cxad">
		  <cx:name>%mlTidning% - %mlKundnr% - %mlKundnamn%</cx:name>
		</cx:folder>
		)
		HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
		OPTS = Charset: utf-8
		HTTPRequest( URL, DATA, HEAD, OPTS )
		; --------------------------------------------------------------------
		gosub, cxSplashOff
		MsgBox,1, Kund inlagd, Kund inlagd. Boka kampanj?
		ifMsgBox, Cancel
		{
			goto, cxDie
		}
		ifMsgBox, OK
		{
			StringSplit, xmlSplit, DATA, >
			StringSplit, xmlSplit, xmlSplit6, <
			kundID = %xmlSplit1% ; kundID innehåller kundens ID
			goto, cxKampanjbokning
		}
		return
	}
	IfMsgBox, No
		goto, cxDie
return

cxKampanjbokning:
	Gui, 1:Destroy ; För säkerhets skull?
	getFormat(mlEnhet)
	if ( format = "")
	{
		Msgbox, Kunde inte hämta produkt, avbryter.
		goto, cxDie
	}
	advertisingFolder = %mlTidning% - %mlKundnr% - %mlKundnamn%
	campaign = %mlTidning% - %format%%mlInternetenhet% - %mlOrdernummer%

	; Sätter rätt format på start och stoppdatum samt varnar om startdatum passerat
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
	if (skin)
	{
		if (skin != "ERROR")
		{
		Gui, 1:Add, Picture, x0 y0 w450 h250 , %mlpDir%\skin\%skin%
		}
	}

	Gui, 1: Add, GroupBox, x12 y10 w420 h180 , Bokningsöversikt
	Gui, 1: Add, Text, x22 y40 w80 h20 , Typ:
	Gui, 1: Add, DropDownList, x142 y40 w120 h20 vType gType R7, Run On Site||Riktad|Plugg|Retarget|CPC
	Gui, 1: Add, Text, x22 y70 w100 h20 , Advertising Folder:
	Gui, 1: Add, Edit, x142 y70 w280 h20 vadvertisingFolder, %advertisingFolder%
	Gui, 1: Add, Text, x22 y100 w100 h20 , Campaign:
	Gui, 1: Add, Edit, x142 y100 w280 h20 vcampaign, %campaign%
	Gui, 1: Add, Text, x22 y130 w110 h20 , Start- och stoppdatum:
	Gui, 1: Add, DateTime, x142 y130 w120 h20 vmlStartdatum Choose%mlStartdatumStrip%00, yyyy-MM-dd  HH:mm
	Gui, 1: Add, Text, x282 y130 w10 h20 , -
	Gui, 1: Add, DateTime, x302 y130 w120 h20 vmlStoppdatum Choose%mlStoppdatumStrip%2359, yyyy-MM-dd  HH:mm
	Gui, 1: Add, Text, x22 y160 w100 h20 , Exponeringar:
	Gui, 1: Add, Edit, x142 y160 w280 h20 vmlExponeringar, %mlExponeringar%
	Gui, 1: Add, Button, x332 y200 w100 h30 gcxAvbryt, Avbryt
	Gui, 1: Add, Button, x222 y200 w100 h30 gcxBokaKampanj Default, OK
	Gui, 1: Color, FFFFFF
	Gui, 1: +ToolWindow +Caption
	Gui, 1: Show, x726 y420 h250 w450, Bokningsöversikt
return

cxAvbryt:
	Gui, 1:Destroy
return

1GuiClose:
	Gui, 1:Destroy
return


Type:
	Gui, 1:Submit, NoHide
	if (Type = "Retarget" || Type = "CPC" || Type = "Plugg")
	{
		GuiControl, Disable, mlExponeringar
	} else {
		GuiControl, Enable, mlExponeringar
	}
	
	
	
return

cxBokaKampanj:
	Gui, 1:Submit
	Gui, 1:Destroy
	cxProduct(format, Type) ; Returnerar productID
	FormatTime, mlStartdatum, %mlStartdatum%, yyyy-MM-dd HH:mm
	FormatTime, mlStoppdatum, %mlStoppdatum%, yyyy-MM-dd HH:mm
	
	StringSplit, mlStartdatum, mlStartdatum, %A_Space%
	startDatum = %mlStartdatum1%
	startTid = %mlStartdatum2%
	
	StringSplit, mlStoppdatum, mlStoppdatum, %A_Space%
	stoppDatum = %mlStoppdatum1%
	stoppTid = %mlStoppdatum2%
	
	if (Type = "Plugg")
		campaign = %mlTidning% - PLUGG - %format% - %mlOrdernummer%
	if (Type = "Reach")
		campaign = %mlTidning% - REACH - %format% - %mlOrdernummer%
	
	gosub, cxSplashOn	
	; --------------------------- HTTP-Request ---------------------------
	URL = https://cxad.cxense.com/api/secure/campaign/%kundID%
	DATA := ""
	HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
	XML =
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
	FILE = %cxDir%\campaign.xml
	refreshFile(XML, FILE)
	OPTS = Upload: %FILE%
	HTTPRequest( URL, DATA, HEAD, OPTS )
	; --------------------------------------------------------------------

	Gosub, cxBokaKontrakt
return

cxBokaKontrakt:
	xmlOut := DATA
	StringSplit, xmlCampSplit, xmlOut, >
	StringSplit, xmlCampSplit, xmlCampSplit6, <
	kampanjID = %xmlCampSplit1% ; kampanjID innehåller kampanjens ID
	
	Gosub, cxCPM
	
	; --------------------------- HTTP-Request ---------------------------
	URL = https://cxad.cxense.com/api/secure/contract/%kampanjID%
	DATA := ""
	XML =
	(
	<?xml version="1.0" encoding="utf-8"?>
	<cx:cpmContract xmlns:cx="http://cxense.com/cxad/api/cxad">
	<cx:startDate>%startDatum%T%startTid%:00.000+01:00</cx:startDate>
	<cx:endDate>%stoppDatum%T%stoppTid%:59.999+01:00</cx:endDate>
	<cx:priority>0.50</cx:priority>
	<cx:requiredImpressions>%mlExponeringar%</cx:requiredImpressions>
	<cx:costPerThousand class="currency" currencyCode="SEK" value="%CPM%"/>
	</cx:cpmContract>
	)
	if (Type = "CPC" or Type = "Retarget" or Type = "Plugg") ; Om det är en CPC-produkt
	{
		XML =
		(
		<?xml version="1.0"?>
		<cx:cpcContract xmlns:cx="http://cxense.com/cxad/api/cxad">
		<cx:startDate>%startDatum%T%startTid%:00.000+01:00</cx:startDate>
		<cx:endDate>%stoppDatum%T%stoppTid%:59.000+02:00</cx:endDate>
		<cx:priority>0.50</cx:priority>
		</cx:cpcContract>
		)
	}
	FILE = %cxDir%\contract.xml
	refreshFile(XML, FILE)
	HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
	OPTS = Upload: %FILE%
	HTTPRequest( URL, DATA, HEAD, OPTS )
	; --------------------------------------------------------------------
	Goto, cxBokaAdvertisement
return

cxBokaAdvertisement:
	StringReplace, mlKundnamn, mlKundnamn,&,,A
	
	; --------------------------- HTTP-Request ---------------------------
	URL = https://cxad.cxense.com/api/secure/ad/%kampanjID%
	DATA := ""
	XML =
	(
	<?xml version="1.0" encoding="UTF-8"?>
	<cx:ad xmlns:cx="http://cxense.com/cxad/api/cxad">
	<cx:name>%mlKundnamn% - %mlStartdatumStripYY%</cx:name>
	</cx:ad>
	)
	FILE = %cxDir%\advertisement.xml
	refreshFile(XML, FILE)
	HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
	OPTS = Upload: %FILE%
	HTTPRequest( URL, DATA, HEAD, OPTS )
	; --------------------------------------------------------------------
	Goto, cxBokaSiteTargeting
return

cxBokaSiteTargeting:
	StringReplace, mlKundnamn, mlKundnamn,&,,A
	targeting = 
	(
		<cx:publisherTarget>
    		<cx:url>http://%mlSite%</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
	)

	; NT FOLKBLADET
	if (mlSite = "nt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://nt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://folkbladet.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; NT FOLKBLADET MOBIL
	if (mlSite = "mobil.nt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.nt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://m.folkbladet.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.nt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.folkbladet.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; NSD KURIREN
	if (mlSite = "nsd.se" ||mlSite = "kuriren.nu")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://nsd.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://kuriren.nu</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; PT
	if (mlSite = "pt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://pitea-tidningen.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; PT MOBIL
	if (mlSite = "m.pitea-tidn.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.pitea-tidningen.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.pitea-tidningen.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; NSD KURIREN MOBIL
	if (mlSite = "mobil.nsd.se" ||mlSite = "mobil.kuriren.nu")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.nsd.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://m.kuriren.nu</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.nsd.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.kuriren.nu</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; CORREN MOBIL
	if (mlSite = "mobil.corren.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.corren.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.corren.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; MVT MOBIL
	if (mlSite = "mobil.mvt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.mvt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.mvt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; VT MOBIL
	if (mlSite = "mobil.vt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.vt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.vt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; UNT MOBIL
	if (mlSite = "mobil.unt.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.unt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.unt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	if (mlSite = "unt.mobil.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.unt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.unt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; UNT Sigtunabygden
	if (mlSite = "sigtunabygden.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://unt.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}


	; HELAGOTLAND MOBIL
	if (mlSite = "m.helagotland.se")
	{
		targeting = 
		(
		<cx:publisherTarget>
    		<cx:url>http://m.helagotland.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
  		<cx:publisherTarget>
    		<cx:url>http://mobil.helagotland.se</cx:url>
    		<cx:targetType>POSITIVE</cx:targetType>
  		</cx:publisherTarget>
		)
	}

	; --------------------------- HTTP-Request ---------------------------
	URL = https://cxad.cxense.com/api/secure/publisherTargeting/%kampanjID%
	DATA := ""
	XML =
	(
<?xml version="1.0"?>
<cx:publisherTargeting xmlns:cx="http://cxense.com/cxad/api/cxad">
  <cx:templateId>%template%</cx:templateId>
  %targeting%
</cx:publisherTargeting>
	)
	FILE = %cxDir%\targeting.xml
	refreshFile(XML, FILE)
	HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
	OPTS = Upload: %FILE%
	HTTPRequest( URL, DATA, HEAD, OPTS )
	; --------------------------------------------------------------------
	gosub, cxSplashOff
	MsgBox,4, Bokning klar, Inbokning klar, öppna i webbläsaren?
		IfMsgBox, Yes
			run, https://cxad.cxense.com/adv/campaign/%kampanjID%/overview
		IfMsgBox, No
			return
	annonsCheck = 1
	gosub, cxSplashOff
return

cxDie:
return

cxGetAdId:
	gosub, cxKundCheck ; returnerar kundID
	if (order = 1) ; bara om föregående returnerar ett kundID
	{
		gosub, cxSplashOn
		; --------------------------- HTTP-Request ---------------------------
		URL = https://cxad.cxense.com/api/secure/campaigns/%kundID%
		DATA := ""
		HEAD = Content-Type: text/xml`nAuthorization: Basic QVBJLlVzZXI6cGFzczEyMw==
		OPTS = Charset: utf-8
		HTTPRequest( URL, DATA, HEAD, OPTS )
		; --------------------------------------------------------------------

		orderCheck = -%A_Space%%mlOrdernummer%
		StringReplace, DATA, DATA, <cx:campaign>, +, A
		xmlPart = 0
		Loop, Parse, DATA, +
		{
			xmlIndex++
			if InStr(A_LoopField, orderCheck)
			{
				xmlPart = %A_LoopField%
				break
			}
		}
		if (xmlPart != 0)
		{
			StringSplit, xmlSplit, xmlPart, >
			StringSplit, xmlSplit, xmlSplit4, <
			campaignID = %xmlSplit1% ; campaignID innehåller kampanjens ID
		}
		if (xmlPart = 0) ; kampanj kunde inte hittas
		{
			gosub, cxSplashOff
			msgbox, Kunde inte hitta ordernummer.
			order = 0
			Return
		}
	}
	gosub, cxSplashOff
return

Rapport:
	order = 1
	bokning = 0
	gosub, cxGetAdId
	if order = 1
		run, http://rapport.ntm-digital.se/advertiser/%kundID%/campaign/%campaignID%/
Return

oppnaRapportFork:
	if (listCount > 1)
	{
		goto, oppnaFleraRapport
	} else {
		goto, Rapport
	}
Return

oppnaCXfork:
	if (listCount > 1)
	{
		goto, oppnaFleraCX
	} else {
		goto, oppnaCX
	}
Return


oppnaCX:
	order = 1
	bokning = 0
	gosub, cxGetAdId
	if order = 1
		run, https://cxad.cxense.com/adv/campaign/%campaignID%/overview
Return


raknaExponeringar:
	;~ gosub, getList
	totExp = 0
	count = 0
	Loop, Parse, getList, `n 
	{
		StringSplit, kolumn, A_LoopField, `t
		totExp := totExp + kolumn%iniExponeringar%
		count++
		if(count = listCount)
		{
			break
		}
	}
	msgbox, Totalt antal begärda exponeringar: %totExp%
return

raknaValda:
	;~ gosub, getList
	MsgBox, %listCount% annonser markerade.
return

rapporteraFel:
	assign(AnvKort)
	status("Undersöks")
	Gui, Add, Text, x12 y10 w180 h20 , Ange vilka fel som finns på ordern:
	Gui, Add, CheckBox, x12 y40 w180 h30 vGUIsen , Sent bokad
	Gui, Add, CheckBox, x12 y70 w180 h30 vGUImanus, Manus saknas/ej komplett
	Gui, Add, CheckBox, x12 y100 w180 h30 vGUImaterial, Material saknas
	Gui, Add, Button, x52 y130 w90 h30 grapporteraFelOK , OK
	Gui, Color, FFFFFF
	Gui, +ToolWindow +Caption
	Gui, Show, xCenter yCenter h166 w208, Felrapport
return

rapporteraFelOK:
	Gui, Submit
	Gui, Destroy
	v1 = 
	v2 = 
	v3 = 
	if GUIsen = 1
		v1 = Sent bokad (material och manus ska finnas 2 arbetsdagar innan startdatum). Om inte ordern bokas fram finns risk att den inte blir producerad.`n
	if GUImanus = 1
		v2 = Manus saknas/är inkomplett i interna noteringar. Interna noteringar ska alltid innehålla manus.`n
	if GUImaterial = 1
		v3 = Material saknas eller har ett startdatum senare än webbannonsens. Material ska vara skickat till digital.material@ntm.eu 2 arbetsdagar innan annonsstart.`n
	mailTo(mlSaljare, "Ej komplett/korrekt order", mlOrdernummer, mlKundnamn)
	Send, Ordern, {Ctrl down}f{Ctrl up}%mlKundnamn% (%mlOrdernummer%){Ctrl down}f{Ctrl up}, är inte komplett/korrekt och kan inte produceras till det datum den är bokad.`nVar vänlig se över bokningen.`n`n{Ctrl down}f{Ctrl up}Upptäckt problem:{Ctrl down}f{Ctrl up}`n%Sentbokad%%Manus%%Material%

	Send, %v1%
	Send, %v2%
	Send, %v3%
return

lager:
	XML = 
	Loop, Parse, getList, `n
	{
		StringSplit, kolumn, A_LoopField, %A_Tab%
		produkt := kolumn%iniProdukt%
		StringSplit, prodArray, produkt , %A_Space%
		Tidning = %prodArray1%
		getFormat(kolumn%iniEnhet%) ; sätter format
		Ordernr = %kolumn1%
		Start := kolumn%iniStart%
		Stopp := kolumn%iniStopp%

		;~ StringTrimRight, startYear, Start, 6 ; startYear
		;~ StringTrimRight, startMonth, Start, 3 ; startMonth
		;~ StringTrimLeft, startMonth, startMonth, 5 ; startMonth
		;~ StringTrimLeft, startDay, Start, 8 ; startDay
		;~ startMonth := startMonth-1
		;~ Start = %startYear%-%startMonth%-%startDay%
		
		;~ StringTrimRight, stoppYear, Stopp, 6 ; stoppYear
		;~ StringTrimRight, stoppMonth, Stopp, 3 ; stoppMonth
		;~ StringTrimLeft, stoppMonth, stoppMonth, 5 ; stoppMonth
		;~ StringTrimLeft, stoppDay, Stopp, 8 ; stoppDay
		;~ stoppMonth := stoppMonth-1
		;~ Stopp = %stoppYear%-%stoppMonth%-%stoppDay%

		Kundnamn := kolumn%iniKundnamn%
		StringReplace, Kundnamn, Kundnamn, & , &amp;, All
		Exponeringar := kolumn%iniExponeringar%
		
		if (Tidning = "NT")
		{
			Tidning = NTFB
		}
		
		addToXML =
		(
		<kampanj>
			<tidning>%Tidning%</tidning>
			<format>%format%</format>
			<ordernr>%Ordernr%</ordernr>
			<kund>%Kundnamn%</kund>
			<start>%Start%</start>
			<stopp>%Stopp%</stopp>
			<exponeringar>%Exponeringar%</exponeringar>
		</kampanj>
		)
		
		XML = %XML%%addToXML%
	}
	fullXML =
	(
	<lager>
	<timestamp>%A_Now%</timestamp>
	%XML%
	</lager>		
	)
	
	FileDelete, %lagerDir%\lager.xml
	FileEncoding, UTF-8
	FileAppend, %fullXML%, %lagerDir%\lager.xml
	FileEncoding
	Msgbox, XML-fil genererad
return




oppnaFleraCX:
	antalAnnonser := listCount
	aktuellAnnons = 1
	Msgbox, 4, Öppna flera annonser, Öppna %listCount% annonser i Cxense?
	IFmsgbox, yes
		{
		while aktuellAnnons <= antalAnnonser
		{
			listRow := getListRow%aktuellAnnons%
			Stringsplit, kolumn, listRow, `t

			mlStartdatum := kolumn%iniStart%
			mlStoppdatum := kolumn%iniStopp%
			mlExponeringar := kolumn%iniExponeringar%
			mlKundnr := kolumn%iniKundnr%
			mlKundnamn := kolumn%iniKundnamn%
			mlSaljare := kolumn%iniSaljare%
			mlProdukt := kolumn%iniProdukt%
			mlOrdernummer = %kolumn1%

			gosub, oppnaCX
			sleep, 500
			aktuellAnnons++
		}
	}
return


oppnaFleraRapport:
	antalAnnonser := listCount
	aktuellAnnons = 1
	Msgbox, 4, Öppna flera annonser, Öppna %listCount% annonser i Rapportverktyget?
	IFmsgbox, yes
		{
		while aktuellAnnons <= antalAnnonser
		{
			listRow := getListRow%aktuellAnnons%
			Stringsplit, kolumn, listRow, `t

			mlStartdatum := kolumn%iniStart%
			mlStoppdatum := kolumn%iniStopp%
			mlExponeringar := kolumn%iniExponeringar%
			mlKundnr := kolumn%iniKundnr%
			mlKundnamn := kolumn%iniKundnamn%
			mlSaljare := kolumn%iniSaljare%
			mlProdukt := kolumn%iniProdukt%
			mlOrdernummer = %kolumn1%

			gosub, Rapport
			sleep, 500
			aktuellAnnons++
		}
	}
return

bokaFlera:
	antalAnnonser := listCount
	aktuellAnnons = 1
	Msgbox, 4, Boka flera annonser, Boka %listCount% annonser i Cxense?
	IFmsgbox, yes
	{
		while aktuellAnnons <= antalAnnonser
		{
			listRow := getListRow%aktuellAnnons%
			Stringsplit, kolumn, listRow, `t

			mlStartdatum := kolumn%iniStart%
			mlStoppdatum := kolumn%iniStopp%
			mlExponeringar := kolumn%iniExponeringar%
			mlKundnr := kolumn%iniKundnr%
			mlKundnamn := kolumn%iniKundnamn%
			mlSaljare := kolumn%iniSaljare%
			mlProdukt := kolumn%iniProdukt%
			mlOrdernummer = %kolumn1%
			mlEnhet := kolumn%iniEnhet%

			StringSplit, prodArray, mlProdukt , %A_Space%
			mlTidning = %prodArray1%
			mlSite = %prodArray2%

			if (mlSite = "gotland.net")
			{
				mlTidning = GN
			}
			if (mlSite = "nt.se")
			{
				mlTidning = NTFB
			}

			annonsCheck = 0
			gosub, cxBokning
			aktuellAnnons++
			while annonsCheck = 0
			{
				sleep, 1000
			}
			
		}
		Sleep, 1000
		Msgbox, %listCount% annonser inbokade!
	}
return



;--------------
;	Statusar
;--------------

statusKlar:
	status("klar")
return

statusNy:
	status("ny")
return

statusBearbetas:
	status("bearbetas")
return

statusVilande:
	status("vilande")
return

statusManusMail:
	status("vilande")
	assign("Manus på Mail")
return

statusKorrSkickat:
	status("korrektur skickat")
return

statusKorrekturKlart:
	status("korrektur klart")
return

statusUndersoks:
	status("undersöks")
return

statusRep:
	status("repetition")
return

statusLevFardig:
	status("Lev. färdig")
return

statusPreflight:
	status("preflight")
return

statusArkiverad:
	status("Arkiverad")
return

statusTilldelaBearbetas:
	status("bearbetas")
	assign(AnvKort)
return

statusBokad:
	status("Bokad")
return

statusObekraftad:
	status("Obekräftad")
return

statusEjkomplett:
	status("Ej komplett manus")
return

; Kopiera kampanjer
copyCampaigns:
	i = 1
	copyList =
	;~ gosub, getList
	loop, parse, getList
	{
		if (i > listCount)
		{	
			break
		}
		thisRow := getListRow%i%
		copyList = %copyList%`n%thisRow%
		i++
	}
	clipboard = %copyList%
	
return


mlNotes:
	ControlGetText, getInterna, Edit3, Atex MediaLink
	FormatTime, Time
	mlNote = 
	(
	
*** REDIGERAD ***
%AnvKort% - %Time%:

	)
	Gui, 10:Add, Edit, w300 r20 vmlNotes, %mlNote%
	Gui, 10:Add, Button, Default gmlNotesSubmit, Spara
	Gui, 10:Show
return

10GuiClose:
	Gui, 10:Destroy
return

mlNotesSubmit:
	Gui, 10:Submit
	Gui, 10:Destroy
	FileDelete, %notesDir%\%mlOrdernummer%.txt
	FileAppend, %mlNotes%, %notesDir%\%mlOrdernummer%.txt
	
		
return
;--------------
;	Tilldela
;--------------

tilldelaMig:
	assign(AnvKort)
return

tilldelaAnnan:
	Send, !a
return

tilldelaIngen:
	assign(" ")
return

; INSTÄLLNINGAR
settings:
if (NoteWin = 1)
{
	noteCheck = Checked
} else {
	noteCheck =
}

if (skin = "ERROR")
{
	skin = 
}
Gui, 5:Add, CheckBox, x12 y10 w180 h30 vnoteWinCheck %noteCheck%, Stort noteringsfönster + cxMini
Gui, 5:Add, Edit, x12 y70 w180 h30 vskin r1, %skin%
Gui, 5:Add, Text, x12 y50 w180 h20, cx-skin: ( filnamn.jpg - tom om inget)
Gui, 5:Add, Button, x52 y120 w100 h30 gnoteReset, Återställ noteringsfönster
Gui, 5:Add, Button, x52 y160 w100 h30 gsettingsSave, Spara
Gui, 5: +ToolWindow +Caption
Gui, 5:Color, FFFFFF
Gui, 5:Show, w207 h200, Inställningar
return

5GuiClose:
Gui, 5:destroy
return

GuiClose:
Gui, destroy
return

settingsSave:
Gui, 5:Submit
IniWrite, %skin%, %mlpDir%\settings.ini, Settings, Skin
IniWrite, %noteWinCheck%, %mlpDir%\settings.ini, Settings, NoteWin
reload
return

noteReset:
	Gui, 5:destroy
	FileDelete, %mlpDir%\notewin.ini
	MsgBox, Noteringsfönstret återställt.
return

citat:
FileDelete, %A_MyDocuments%\MedialinkPlus\citat.txt
UrlDownloadToFile, http://citatet.se, %A_MyDocuments%\MedialinkPlus\citat.txt
Sleep, 1000
FileRead, citatBlob, *P65001 %A_MyDocuments%\MedialinkPlus\citat.txt
StringReplace, citatBlob, citatBlob, &#8221`;, ”, A
StringReplace, citatBlob, citatBlob, ”, £, A
StringReplace, citatBlob, citatBlob, &#8211`;, -, A

StringSplit, citatArray, citatBlob , £
citat = %citatArray2%

StringReplace, citatBlob, citatBlob, <span>, €, A
StringReplace, citatBlob, citatBlob, </span>, €, A
StringSplit, citatArray, citatBlob , €
author = %citatArray2%
return

setCitat:
	mlNewTitle = %mlTitle%     ::     %citat% - %author%
	WinSetTitle, Atex MediaLink,,%mlNewTitle%
return

cycleCitat:
	gosub, citat
	gosub, setCitat
return

^!k::
	gosub, getAnvnamn
	msgbox, %Anvandare%
return

#include cxmini.ahk
#include weblink.ahk
FileEncoding, UTF-8
^RButton::
	MouseGetPos, , , id, control
	IfInString, control, SysListView
	{
		SLV = 1
	} else {
		SLV = 0
	}
WinGetTitle, Windowtext, Atex MediaLink
	StringSplit, WindowSplit, Windowtext, =
	Anvandare =  %WindowSplit2%

if (Anvandare = "martinve")
		{
			Start = %listArr9%
			Stopp = %listArr10%
			Exponeringar = %listArr13%
			Kundnr = %listArr7%
			Kundnamn = %listArr8%
			Ordernr = %listArr1%
		}
	if (Anvandare = "anniejo")
		{
			Start = %listArr5%
			Stopp = %listArr6%
			Exponeringar = %listArr11%
			Kundnr = %listArr10%
			Kundnamn = %listArr8%
			Ordernr = %listArr1%
		}
	if (Anvandare = "annicasv")
		{
			Start = %listArr5%
			Stopp = %listArr6%
			Exponeringar = %listArr9%
			Kundnr = %listArr10%
			Kundnamn = %listArr11%
			Ordernr = %listArr1%
		}
	if (Anvandare = "mikaellu")
		{
			Start = %listArr8%
			Stopp = %listArr9%
			Exponeringar = %listArr12%
			Kundnr = %listArr13%
			Kundnamn = %listArr5%
			Ordernr = %listArr1%
		}
	if (Anvandare = "dennisst")
		{
			StartRAW = kolumn6
			StoppRAW = kolumn7
			ExponeringarRAW = kolumn10
			KundnrRAW = kolumn11
			KundnamnRAW = kolumn12
			OrdernrRAW = kolumn1
			TidningRAW = kolumn2
			ProduktRAW = kolumn5
		}

	If ((WinActive("ahk_class wxWindowClassNR") and SLV = 1))
	{
		menu, LagerMenu, add, Lager, Lager
		menu, LagerMenu, show
	}
	return


Lager:
	list = 
	ControlGet, allItems, List, , %control%, Atex MediaLink
	Loop, Parse, allItems, `n
	{
		StringSplit, kolumn, A_LoopField, %A_Tab%
		
		;TIDNING
		TidningKol := %TidningRAW%
		StringSplit, produkt, TidningKol, %A_Space%

		Tidning = %produkt1%
		if (produkt2 = "gotland.net")
		{
			Tidning = GN
		}
		if (produkt2 = "nt.se" || produkt2 = "mobil.nt.se")
		{
			Tidning = NTFB
		}


		;FORMAT
		Produkt := %ProduktRAW%

		if (Produkt = "Guld" || Produkt = "Widescreen" || Produkt = "Guld sidfot" || Produkt = "Widescreen 240" )
		{
			Format = WID
		}
		if (Produkt = "Mittbanner" || Produkt = "Mittbanner 1" ||Produkt = "Modul" ||Produkt = "Modul 1" ||Produkt = "Modul 2")
		{
			Format = MOD
		}
		if (Produkt = "Hissen" || Produkt = "Hiss" ||Produkt = "Skyskrapa" ||Produkt = "Outsider")
		{
			Format = OUT
		}
		if (Produkt = "Bigbanner" || Produkt = "Panorama" || Produkt = "Panorama XL" || Produkt = "Bigbanner XL" || Produkt = "Panorama 120" || Produkt = "Panorama 240")
		{
			Format = PAN
		}
		if (Produkt = "Mobil Bottom Panorama XL" || Produkt = "Mobil Top Panorama" ||Produkt = "Mobil Top Panorama XL" ||Produkt = "Mobil Top Takeover")
		{
			Format = MOB
		}
		if (Produkt = "Reach 250" || Produkt = "Reach 468")
		{
			Format = REACH
		}
		if (Produkt = "Textannons")
		{
			Format = TXT
		}
		if (Produkt = "Kvadrat")
		{
			Format = 180
		}
		if (Produkt = "Stortavla")
		{
			Format = 380
		}
		if (Produkt = "Tema" || Produkt = "Temabanner")
		{
			Format = TEMA
		}
		if (Produkt = "Toppjobb")
		{
			Format = TOPPJOBB
		}
		if (Produkt = "Väderspons")
		{
			Format = VADER
		}
		if (Produkt = "Väderspons")
		{
			Format = VADER
		}
		if (Produkt = "Affärsliv Guld" || Produkt = "Affärsliv Widescreen")
		{
			Format = AFWID
		}
		if (Produkt = "Eurosize")
		{
			Format = 1080
		}
		if (Produkt = "Hamnbron")
		{
			Format = 1024
		}
		if (Produkt = "Sjötull")
		{
			Format = 512
		}
		if (Produkt = "Skylt HD")
		{
			Format = 1920
		}
		if (Produkt = "Ståthögaleden")
		{
			Format = 768
		}
		if (Produkt = "Söderköpingsvägen")
		{
			Format = 1024
		}

		; ORDERNUMMER
		Ordernr := %OrdernrRAW%

		; START / STOPPDATUM
		Start := %StartRAW%
		Stopp := %StoppRAW%

		; KUNDNAMN
		Kundnamn := %KundnamnRAW%
		StringReplace, Kundnamn, Kundnamn, & , &amp;, All

		; EXPONERINGAR
		Exponeringar := %ExponeringarRAW%



		addToList = 
		(
	<kampanj>
		<tidning>%Tidning%</tidning>
		<format>%Format%</format>
		<ordernr>%Ordernr%</ordernr>
		<kund>%Kundnamn%</kund>
		<start>%Start%</start>
		<stopp>%Stopp%</stopp>
		<exponeringar>%Exponeringar%</exponeringar>
	</kampanj>

		)

		list = %list%%addToList%
	}

	xml = 
	(
	<lager>
		<timestamp>%A_Now%</timestamp>
	%list%
	</lager>
	)

	lagerDir = X:\digital.ntm.eu\lager
	;lagerDir = C:\AHK
	FileDelete, %lagerDir%\lager.xml
	FileAppend, %xml%, %lagerDir%\lager.xml
	Msgbox, XML-fil genererad
	Reload
	Sleep, 1000
	return
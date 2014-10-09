FileEncoding, UTF-8

^RButton::
	MouseGetPos, , , id, control
	IfInString, control, SysListView
	{
		SLV = 1
	} else {
		SLV = 0
	}
	If ((WinActive("ahk_class wxWindowClassNR") and SLV = 1))
	{
		menu, LagerMenu, add, Lager, Lager
		menu, LagerMenu, show
	}
	return

Lager:
	; RESET
	W = 0
	O = 0
	M = 0
	P = 0

	cW = 0
	cO = 0
	cM = 0
	cP = 0

	ntW = 0
	ntO = 0
	ntM = 0
	ntP = 0

	vW = 0
	vO = 0
	vM = 0
	vP = 0

	mW = 0
	mO = 0
	mM = 0
	mP = 0

	gW = 0
	gO = 0
	gM = 0
	gP = 0

	uW = 0
	uO = 0
	uM = 0
	uP = 0

	pW = 0
	pO = 0
	pM = 0
	pP = 0

	nmW = 0
	nmO = 0
	nmM = 0
	nmP = 0


	i = 0

	; Hämta lista
	ControlGet, listLoop, List, Selected, SysListView321, Atex MediaLink ; Returnerar alla markerade enheter
	ControlGet, listCount, List, Count Selected, SysListView321, Atex MediaLink ; Returnerar antalet markerade enheter

	; Loopa
	Loop, Parse, listLoop, `n
	{
		i++
		StringSplit, kolumn, A_LoopField, `t

		if (kolumn5 = "Guld" || kolumn5 = "Guld sidfot" || kolumn5 = "Widescreen")
		{
			IfInString, kolumn2, corren
			{
				cW := cW + kolumn10
				cWcount++
				kundnamn = %kolumn12%
				cWAd%cWcount% = 
				(
				<campaign>
					<name>CO - WID - %kundnamn%</name>
					<startdate>%kolumn6%</startdate>
					<enddate>%kolumn7%</enddate>
					<impressions>%kolumn10%</impressions>
				</campaign>
				)
			}

			IfInString, kolumn2, nt.se
				ntW := ntW + kolumn10
				

			IfInString, kolumn2, UNT
				uW := uW + kolumn10
				

			IfInString, kolumn2, helagotland.se
				gW := gW + kolumn10
				

			IfInString, kolumn2, mvt.se
				mW := mW + kolumn10
				

			IfInString, kolumn2, VT vt.se
				vW := vW + kolumn10
				

			IfInString, kolumn2, kuriren.nu
				nmW := nmW + kolumn10
				

			IfInString, kolumn2, nsd.se
				nmW := nmW + kolumn10
				

			IfInString, kolumn2, pt.se
				pW := pW + kolumn10
				
		}

		if (kolumn5 = "Modul" || kolumn5 = "Mittbanner 1")
		{
			IfInString, kolumn2, corren
				cM := cM + kolumn10
				

			IfInString, kolumn2, nt.se
				ntM := ntM + kolumn10
				

			IfInString, kolumn2, UNT
				uM := uM + kolumn10
				

			IfInString, kolumn2, helagotland.se
				gM := gM + kolumn10
				

			IfInString, kolumn2, mvt.se
				mM := mM + kolumn10
				

			IfInString, kolumn2, VT vt.se
				vM := vM + kolumn10
				

			IfInString, kolumn2, kuriren.nu
				nmM := nmM + kolumn10
				

			IfInString, kolumn2, nsd.se
				nmM := nmM + kolumn10
				

			IfInString, kolumn2, pt.se
				pM := pM + kolumn10
				
		}

		if (kolumn5 = "Hissen" || kolumn5 = "Outsider" || kolumn5 = "Skyskrapa")
		{
			IfInString, kolumn2, corren
				cO := cO + kolumn10
				

			IfInString, kolumn2, nt.se
				ntO := ntO + kolumn10
				

			IfInString, kolumn2, UNT
				uO := uO + kolumn10
				

			IfInString, kolumn2, helagotland.se
				gO := gO + kolumn10
				

			IfInString, kolumn2, mvt.se
				mO := mO + kolumn10
				

			IfInString, kolumn2, VT vt.se
				vO := vO + kolumn10
				

			IfInString, kolumn2, kuriren.nu
				nmO := nmO + kolumn10
				

			IfInString, kolumn2, nsd.se
				nmO := nmO + kolumn10
				

			IfInString, kolumn2, pt.se
				pO := pO + kolumn10
				
		}

		if (kolumn5 = "Panorama 1" || kolumn5 = "Panorama 2" || kolumn5 = "Bigbanner" || kolumn5 = "Bigbanner XL" || kolumn5 = "Panorama" || kolumn5 = "Panorama 1 (980x240)" || kolumn5 = "Panorama 1 (980x120)")
		{
			IfInString, kolumn2, corren
				cP := cP + kolumn10
				

			IfInString, kolumn2, nt.se
				ntP := ntP + kolumn10
				

			IfInString, kolumn2, UNT
				uP := uP + kolumn10
				

			IfInString, kolumn2, helagotland.se
				gP := gP + kolumn10
				

			IfInString, kolumn2, mvt.se
				mP := mP + kolumn10
				

			IfInString, kolumn2, VT vt.se
				vP := vP + kolumn10
				

			IfInString, kolumn2, kuriren.nu
				nmP := nmP + kolumn10
				

			IfInString, kolumn2, nsd.se
				nmP := nmP + kolumn10
				

			IfInString, kolumn2, pt.se
				pP := pP + kolumn10
				
		}

		if (i = listCount)
		{
			break
		}

	}
	Total = 
	(
NTFB
	Wide:		%ntW%
	Outsider:		%ntO%
	Modul:		%ntM%
	Panorama:	%ntP%

CO
	Wide:		%cW%
	Outsider:		%cO%
	Modul:		%cM%
	Panorama:	%cP%

HG
	Wide:		%gW%
	Outsider:		%gO%
	Modul:		%gM%
	Panorama:	%gP%

MVT
	Wide:		%mW%
	Outsider:		%mO%
	Modul:		%mM%
	Panorama:	%mP%

VT
	Wide:		%vW%
	Outsider:		%vO%
	Modul:		%vM%
	Panorama:	%vP%

NM
	Wide:		%nmW%
	Outsider:		%nmO%
	Modul:		%nmM%
	Panorama:	%nmP%

UNT
	Wide:		%uW%
	Outsider:		%uO%
	Modul:		%uM%
	Panorama:	%uP%

PT
	Wide:		%pW%
	Outsider:		%pO%
	Modul:		%pM%
	Panorama:	%pP%
	)
	Msgbox, %Total%


; Corren Wide
Loop %cWcount%
{
	index = cWAd%A_Index%
	StringReplace, index, index, & , &amp;, All]
	cWlist := cWlist . cWAd%A_Index%
}


xml = 
(	<?xml version="1.0" encoding="UTF-8"?>
	<lager>
		<item>
			<site>NTFB</site>
			<wid>%ntW%</wid>
			<out>%ntO%</out>	
			<mod>%ntM%</mod>
			<pan>%ntP%</pan>
		</item>
		<item>
			<site>CO</site>
			<wid>%cW%</wid>
			<out>%cO%</out>	
			<mod>%cM%</mod>
			<pan>%cP%</pan>
		</item>
		<item>
			<site>HG</site>
			<wid>%gW%</wid>
			<out>%gO%</out>	
			<mod>%gM%</mod>
			<pan>%gP%</pan>
		</item>
		<item>
			<site>MVT</site>
			<wid>%mW%</wid>
			<out>%mO%</out>	
			<mod>%mM%</mod>
			<pan>%mP%</pan>
		</item>
		<item>
			<site>VT</site>
			<wid>%vW%</wid>
			<out>%vO%</out>	
			<mod>%vM%</mod>
			<pan>%vP%</pan>
		</item>
		<item>
			<site>NM</site>
			<wid>%nmW%</wid>
			<out>%nmO%</out>	
			<mod>%nmM%</mod>
			<pan>%nmP%</pan>
		</item>
		<item>
			<site>UNT</site>
			<wid>%uW%</wid>
			<out>%uO%</out>	
			<mod>%uM%</mod>
			<pan>%uP%</pan>
		</item>
		<item>
			<site>PT</site>
			<wid>%pW%</wid>
			<out>%pO%</out>	
			<mod>%pM%</mod>
			<pan>%pP%</pan>
		</item>
		<timestamp>%A_Now%</timestamp>

	<test>
		%cWlist%
	</test>

	</lager>
)

;lagerDir = X:\digital.ntm.eu\lager
lagerDir = C:\AHK
FileDelete, %lagerDir%\lager.xml
FileAppend, %xml%, %lagerDir%\lager.xml
Msgbox, XML-fil genererad
Reload
Sleep, 1000
return
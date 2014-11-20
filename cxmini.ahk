^!+r::
Reload
Return

^!c::
	gosub, cxMini
return

cxMini:
listVy = adFolder
allAds = 
GuiControl,8: ,adList, %allAds%
Gui, 8:Add, ListBox, x5 y90 w360 h280 vadList gadList, 
Gui, 8:Add, Button, x5 y10 w360 h40 gUpdate, Visa kunder
Gui, 8:Add, Edit, x5 y60 w360 h20 gSearch vSearch
Gui, 8:Color, FFFFFF
Gui, 8: +ToolWindow +Caption
Gui, 8:Show, w369 h370, nonXense
GuiControl, 8: Focus, Search
Gosub, Update
return

8GuiClose:
Gui, 8:Destroy
allAds = 
return

Update:
WinActivate, nonXense
GuiControl,8: ,Search, 
GuiControl, 8: Focus, Search
listVy = adFolder
FileDelete, %cxDir%getAds.xml
FileDelete, %cxDir%getAds.bat
call = https://cxad.cxense.com/api/secure/folder/advertising
bat = 
(
G:
cd G:\NTM\NTM Digital Produktion\cURL\bin
curl -s -H "Content-type: text/xml" -u API.User:pass123 -X GET %call% > %cxDir%getAds.xml
)
FileEncoding, UTF-8-RAW
FileAppend, %bat%, %cxDir%getAds.bat
FileEncoding
Run, %cxDir%getAds.bat,,Min


Sleep, 100
WinWaitClose, C:\Windows\system32\cmd.exe
FileEncoding, UTF-8-RAW
FileRead, getAds, %cxDir%getAds.xml
FileEncoding
StringReplace, getAds, getAds, <cx:name>, ~, All

allAds = 
loop, parse, getAds, ~
{
	StringSplit, kund, A_LoopField, <
	allAds = %allAds%|%kund1%|
}
GuiControl, 8:-Redraw, adList
GuiControl, 8:,adList, %allAds%
GuiControl, 8:+Redraw, adList
return

adList:
	if A_GuiControlEvent <> DoubleClick
		return

	if A_GuiControlEvent = DoubleClick
		gosub, adFolders
		return
	return


Search:
if (listVy = "adFolder")
{
	GuiControlGet, Search
	GuiControl, 8:-Redraw, adList
	allAds = 
	GuiControl,8: ,adList, %allAds%
	loop, parse, getAds, ~
	{
		StringSplit, kund, A_LoopField, <
		IfInString, kund1, %Search%
			allAds = %allAds%|%kund1%|
	}
	GuiControl,8: ,adList, %allAds%
	GuiControl, 8: +Redraw, adList
}

if (listVy = "Campaigns")
{
	GuiControlGet, Search
	GuiControl, 8: -Redraw, adList
	allCampaigns =
	GuiControl,8: ,adList, %allCampaigns%
	loop, parse, getCampaigns, ~
	{
		StringSplit, kampanj, A_LoopField, <
		IfInString, kampanj1, %Search%
			allCampaigns = %allCampaigns%|%kampanj1%|
	}
	GuiControl,8: ,adList, %allCampaigns%
	GuiControl,8: +Redraw, adList
}
return

Contract:

return

adFolders:
WinActivate, nonXense
GuiControl, ,8:Search, 
GuiControl, Focus, 8:Search
if (listVy = "Campaigns")
{
	goto, Campaigns
}
StringReplace, getAds, getAds, <cx:childFolder, ¨, All
GuiControlGet, adList, 8:
StringSplit, adSplit, adList, -
StringReplace, kundnr, adSplit2, %A_Space%,,A

Loop, Parse, getAds, ¨
{
	if InStr(A_LoopField, kundnr)
	{
		kundBlock = %A_LoopField%
		break
	}
}

StringSplit, kund, kundBlock, >
StringSplit, folderID, kund5, <
folderID := folderID1 ; folderID inneh?ller IDt för den kund man dubbelklickade på
StringSplit, kundnamn, kund6, <
StringTrimLeft, kundNamn, kundnamn1, 1 ; kundNamn innehåller namnet på den kund man dubbelklickade på
FileDelete, %cxDir%getCampaigns.bat
FileDelete, %cxDir%getCampaigns.xml
call = https://cxad.cxense.com/api/secure/campaigns/
bat = 
(
G:
cd G:\NTM\NTM Digital Produktion\cURL\bin
curl -s -H "Content-type: text/xml" -u API.User:pass123 -X GET %call%%folderID% > %cxDir%getCampaigns.xml
)
FileEncoding, UTF-8-RAW
FileAppend, %bat%, %cxDir%getCampaigns.bat
FileEncoding
Run, %cxDir%getCampaigns.bat,,Min
Sleep, 100

WinWaitClose, C:\Windows\system32\cmd.exe
FileEncoding, UTF-8-RAW
FileRead, getCampaigns, %cxDir%getCampaigns.xml
FileEncoding
StringReplace, getCampaigns, getCampaigns, <cx:name>, ~, All

allCampaigns = 
loop, parse, getCampaigns, ~
{
	StringSplit, kampanj, A_LoopField, <
	allCampaigns = %allCampaigns%|%kampanj1%|
}
StringTrimLeft, allCampaigns, allCampaigns, 2
GuiControl,8: -Redraw, adList
GuiControl,8: ,adList, %allCampaigns%
GuiControl, 8: +Redraw, adList
listVy = Campaigns
return

Campaigns:
GuiControlGet, adList , 8:
kampanjNamn = %adlist%
FileEncoding, UTF-8-RAW
FileRead, getCampaigns, %cxDir%getCampaigns.xml
FileEncoding

StringReplace, getCampaigns, getCampaigns, <cx:campaignID>, ~, All
Loop, Parse, getCampaigns, ~
{
	if InStr(A_LoopField, kampanjNamn)
	{
		campaignBlock = %A_LoopField%
		break
	}
}
StringSplit, campaignID, campaignBlock , <
campaignID = %campaignID1%

run, https://cxad.cxense.com/adv/campaign/%campaignID%/overview
Gui, 8:Destroy
;FileDelete, %cxDir%getContractList.bat
;FileDelete, %cxDir%getContractList.xml
;call = https://cxad.cxense.com/api/secure/contracts/
;bat = 
;(
;G:
;cd G:\NTM\NTM Digital Produktion\cURL\bin
;curl -s -H "Content-type: text/xml" -u API.User:pass123 -X GET %call%%campaignID% > %cxDir%getContractList.xml
;)
;FileEncoding, UTF-8-RAW
;FileAppend, %bat%, %cxDir%getContractList.bat
;FileEncoding
;Run, %cxDir%getContractList.bat,,Min
;Sleep, 100
;WinWaitClose, C:\Windows\system32\cmd.exe
;FileEncoding, UTF-8-RAW
;FileRead, getContractList, %cxDir%getContractList.xml
;FileEncoding

;StringReplace, getContractStartDate, getContractList, <cx:startDate>, ~, All
;StringSplit, getContractStartDate, getContractStartDate , ~, 
;StringSplit, getContractStartDate, getContractStartDate2 , <, 
;StringSplit, getContractStartDate, getContractStartDate1 , T, 
;startdatum = %getContractStartDate1%

;StringReplace, getContractStopDate, getContractList, <cx:endDate>, ~, All
;StringSplit, getContractStopDate, getContractStopDate , ~, 
;StringSplit, getContractStopDate, getContractStopDate2 , <, 
;StringSplit, getContractStopDate, getContractStopDate1 , T, 
;slutdatum = %getContractStopDate1%

;StringReplace, getImpressions, getContractList, <cx:requiredImpressions>, ~, All
;StringSplit, getImpressions, getImpressions , ~, 
;StringSplit, getImpressions, getImpressions2 , <, 
;exponeringar = %getImpressions1%

;Gui, 2:Add, GroupBox, x12 y10 w430 h100 , %kundNamn% / %kampanjNamn%
;Gui, 2:Add, Text, x22 y30 w90 h20 , Startdatum:
;Gui, 2:Add, Text, x122 y30 w90 h20 , %startdatum%
;Gui, 2:Add, Text, x232 y30 w90 h20 , Slutdatum:
;Gui, 2:Add, Text, x332 y30 w100 h20 , %slutdatum%
;Gui, 2:Add, Text, x22 y60 w90 h20 , Exponeringar:
;Gui, 2:Add, Text, x122 y60 w90 h20 , %exponeringar%

;Gui, 2:Show, x285 y107 h123 w458, Kampanjsummering

Return

~Down::
	GuiControlGet, focus, 8:FocusV
IfWinActive, nonXense
	if (focus = "Search")
	{
		GuiControl, 8:Focus, adList
		Send, {Home}
	}
return

~Enter::
	GuiControlGet, focus, 8:FocusV
IfWinActive, nonXense
	if (focus = "adList")
	{
		Gosub, adFolders
	}
return

~Esc::
IfWinActive, nonXense
	Gui, 8:Destroy
	GuiControl, ,adList, 
	return
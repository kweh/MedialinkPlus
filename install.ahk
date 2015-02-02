SetTitleMatchMode, 2
DetectHiddenText, On
DetectHiddenWindows, On
IfNotExist, %A_MyDocuments%\MedialinkPlus
	FileCreateDir, %A_MyDocuments%\MedialinkPlus

UrlDownloadToFile, http://dnns.se/mlp/mlp/latest/MedialinkPlus.exe, %A_MyDocuments%\MedialinkPlus\MedialinkPlus_new.exe
If (ErrorLevel != 1) ; Om filen laddades ner
{
	WinClose, MedialinkPlus
	Msgbox, ,Senaste versionen nedladdad, Senaste versionen av Medialink Plus laddades ner utan problem.
	FileDelete, %A_MyDocuments%\MedialinkPlus\MedialinkPlus.exe ; Ta bort den gamla
	FileMove, %A_MyDocuments%\MedialinkPlus\MedialinkPlus_new.exe, %A_MyDocuments%\MedialinkPlus\MedialinkPlus.exe ; Byt namn på nedladdad fil
	IfNotExist, %A_Desktop%\MedialinkPlus.lnk
	{
		Msgbox, 4, Genväg?, Vill du lägga en genväg på skrivbordet?
		IfMsgBox, Yes
			FileCreateShortcut, %A_MyDocuments%\MedialinkPlus\MedialinkPlus.exe, %A_Desktop%\MedialinkPlus.lnk
	}
	IfNotExist, %A_Startup%\MedialinkPlus.lnk
	{
		Msgbox, 4, Autostart?, Vill du lägga till Medialink Plus i autostart??
		IfMsgBox, Yes
			FileCreateShortcut, %A_MyDocuments%\MedialinkPlus\MedialinkPlus.exe, %A_Desktop%\MedialinkPlus.lnk
	}
	Run, %A_MyDocuments%\MedialinkPlus\MedialinkPlus.exe
}
If (ErrorLevel = 1)
{
	Msgbox, Något gick fel, försök igen.
}

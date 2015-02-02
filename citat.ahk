FileDelete, %A_MyDocuments%\MedialinkPlus\citat.txt
UrlDownloadToFile, http://citatet.se, %A_MyDocuments%\MedialinkPlus\citat.txt
Sleep, 1000
FileRead, citatBlob, *P65001 %A_MyDocuments%\MedialinkPlus\citat.txt
StringReplace, citatBlob, citatBlob, &#8221`;, £, A
StringSplit, citatArray, citatBlob , £
citat = %citatArray3%
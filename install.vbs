dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://raw.githubusercontent.com/DHemken97/GAH_Silent_Install/main/init.bat", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "C:\temp1\init.bat", 2 '//overwrite
end with
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\temp1\init.bat" & Chr(34), 0
Set WshShell = Nothing
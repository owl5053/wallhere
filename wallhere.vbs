'On error resume next 'comment it for debug

set FSO=CreateObject ("Scripting.FileSystemObject")
wallherefile = fso.GetSpecialFolder(2): if right(wallherefile,1)<>"\" then wallherefile=wallherefile & "\" : wallherefile = wallherefile & "wallhere.jpg"

searchstring="natural"

'=== part 1 ====
sUrlRequest = "https://wallhere.com/en/random?q=" & searchstring & "&direction=horizontal"
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
httpfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing

beg=instr(lcase(httpfile),"/en/wallpaper/")
ef=instr(lcase(httpfile),"current-item-photo")
lnk=mid(httpfile,beg,ef-beg-9)
url="https://wallhere.com" & lnk


'=== part 2 ====
sUrlRequest = url
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
httpfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing

beg=instr(lcase(httpfile),"https://get.wallhere.com/photo/")
ef=instr(lcase(httpfile),"current-page-photo")
url=mid(httpfile,beg,ef-beg-9)

Set oXMLHTTP2 = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP2.Open "GET", url, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile wallherefile, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
'objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", wallherefile, "REG_SZ"
'objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False

'use irfanview if you want
objWshShell.Run "c:\Programs\IrfanView\i_view64.exe """ & wallherefile & """ /wall=0 /killmesoftly", 1, False 


Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing
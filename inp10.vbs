Dim config
config=1
inp10 =inputbox("inp10","INP10")
if UBound(Split("readme", inp10))>0 then
CreateObject("Shell.Application").ShellExecute "reader.bat",,, "runas", config
elseif UBound(Split("ipconfig", inp10))>0 then
CreateObject("Shell.Application").ShellExecute "ipconfig.bat",,, "runas", config
elseif UBound(Split("netsh wlan show profiles", inp10))>0 then
CreateObject("Shell.Application").ShellExecute "netsh wlan show profiles.bat",,, "runas", config
elseif UBound(Split("ping", inp10))>0 then
CreateObject("Shell.Application").ShellExecute "ping.bat",,, "runas", config
elseif UBound(Split("tasklist", inp10))>0 then
CreateObject("Shell.Application").ShellExecute "tasklist.bat",,, "runas", config
elseif inp10 <> "" Then
Set WshShell = CreateObject("WScript.Shell")
inp10=Replace(inp10," ","")
inp10=Replace(inp10,"/","")
inp100=inp10
inp10=inp10+" pause"
WshShell.Run inp10
set objA = WshShell.Exec(inp100)
Set fs = CreateObject("Scripting.FileSystemObject")
Set tst = fs.CreateTextFile(inp100+".txt", True)
tst.WriteLine(objA.StdOut.ReadAll())
tst.Close
else
end if
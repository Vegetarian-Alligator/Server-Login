Dim sshtunnel
Dim desktop

Input = InputBox("Enter your Password.  Don't let anyone look over your shoulder!") 
Set tunnel = WScript.CreateObject("WScript.Shell")
Return = tunnel.run("custputty -ssh miner@64.251.26.108 -L 9000:localhost:5901 -pw " & Input,0,false)
Set desktop = WScript.CreateObject("WScript.Shell")
Return = desktop.run("tvnviewer -host=localhost -port=9000",0,true)
Return = tunnel.run("taskkill /im custputty.exe /f",0,false)
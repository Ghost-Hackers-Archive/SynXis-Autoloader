WScript.Sleep 1000
    MsgBox "Click [Okay] to load SynXis PM. It will automatically load itself, no action is necessary.", 10, "SynXis PM Auto Loader"

Set x=CreateObject("WScript.Shell")
    x.Run "https://whg.sabrehospitality.com/Loader/Loader.application?1=1&Lang=en-us"

WScript.Sleep 3000																		
    x.SendKeys "{LEFT}"

WScript.Sleep 500
    x.SendKeys "{ENTER}"

WScript.Sleep 3000																		
    CreateObject("WScript.Shell").Run "TaskKill /F /IM msedge.exe"

WScript.Sleep 1500
    x.AppActivate "Application Run - Security Warning"

WScript.Sleep 2000
    x.SendKeys "%R"

WScript.Sleep 20000
    x.SendKeys "12654"

WScript.Sleep 1000
    x.SendKeys "{TAB}"

WScript.Sleep 1000
    MsgBox "SynXis has been successfully loaded. You can now login.", 0, "SynXis PM Auto Loader"
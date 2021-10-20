' --- Code  Information ---
  ' Code Owner:       Ghost
  ' GitHub:           https://github.com/Ghostridr
  ' Version:          1.0.0
  ' Creation Date:    October 19, 2021
'
' --- Designed for Information ---
  ' Platform:         SynXis
  ' Where to Get:     https://www.sabrehospitality.com/
'
' --- Companies Optimized For  ---
  ' Wynham Hotels
    ' Super 8 
     ' El Dorado, Arkansas
'
' --- The Purpose ---         
  ' To automate the launching process for SynXis PM - making it easier for employee's to continue working.
  ' To free up time normally used to go through the steps - thus increasing productivity since they will no longer have to do this manually.
'
WScript.Sleep 1000
    MsgBox "Click [Okay] to load SynXis PM. It will automatically load itself, no action will be necessary. Please wait for the next popup.", 10, "SynXis Auto Loader"

Set x=CreateObject("WScript.Shell")
    x.Run "https://whg.sabrehospitality.com/Loader/Loader.application?1=1&Lang=en-us"

WScript.Sleep 2000
    x.SendKeys "{LEFT}"

WScript.Sleep 500
    x.SendKeys "{ENTER}"

WScript.Sleep 2000
    x.Run "TaskKill /F /IM msedge.exe"

WScript.Sleep 1500
    x.AppActivate "Application Run - Security Warning"

WScript.Sleep 500
    x.SendKeys "%R"

WScript.Sleep 15000
    x.SendKeys "12654"

WScript.Sleep 500
    x.SendKeys "{TAB}"

WScript.Sleep 500
    MsgBox "SynXis has been successfully loaded. You can now login.", 0, "SynXis PM Auto Loader"
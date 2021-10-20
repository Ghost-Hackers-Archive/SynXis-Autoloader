' --- Code  Information ---
  ' Code Owner:       Ghost Codes
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
    ' LaQuinta Inn
      ' El Dorado, Arkansas
'
' --- The Purpose ---         
  ' To automate the launching process for SynXis PM - making it easier for employee's to continue working.
  ' To free up time normally used to go through the steps - thus increasing productivity since they will no longer have to do this manually.
'
' Instruction box before code is ran.
  WScript.Sleep 1000
    MsgBox "Click [Okay] to load SynXis PM. It will automatically load itself, no action will be necessary. Please wait for the next popup.", 10, "SynXis Auto Loader"
' Changes "WScript.Shell" to "x" and then opens the link that's hidden in the launch button of the site.
    ' This opens a new tab named "Untitled".
  Set x=CreateObject("WScript.Shell")
    x.Run "msedge https://whg.sabrehospitality.com/Loader/Loader.application?1=1&Lang=en-us"
' Moves cursor onto "Open" and presses it.
  WScript.Sleep 2000
    x.SendKeys "{LEFT}"
  WScript.Sleep 500
    x.SendKeys "{ENTER}"
' Closes the web browser now that it's not needed. 
    'Is it possible to close only the tab that is opened in MS Edge?
  WScript.Sleep 2000
    x.Run "TaskKill /F /IM msedge.exe"
' Ensures the security warning popup is focused, selects and presses "Run".
  WScript.Sleep 1500
    x.AppActivate "Application Run - Security Warning"
  WScript.Sleep 500
    x.SendKeys "%R"
' Enters a set of numbers and tabs to a new line.
  WScript.Sleep 15000
    x.SendKeys "12654"
  WScript.Sleep 500
    x.SendKeys "{TAB}"
' Final popup to inform that the script has finished.
  WScript.Sleep 500
    MsgBox "SynXis has been successfully loaded. You can now login.", 0, "SynXis PM Auto Loader"
' SynXis should now be open, ready to use, and the browser closed. Employee's are able to login now.
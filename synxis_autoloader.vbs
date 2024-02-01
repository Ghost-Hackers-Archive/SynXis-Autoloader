' --- Code  Information ---
  ' Code Owner:       Ghost Hackers
    ' GitHub:         [Ghost Hackers](https://github.com/Ghost-Hackers)
  ' Version:          2.0.0
    ' Creation Date:  October 19, 2021
    ' Last Updated:   February 01, 2024
' --- Designed for Information ---
  ' Platform:         SynXis
    ' Where to Get:   [Sabre Hospitality](https://www.sabrehospitality.com/)

' --- Companies Optimized For  ---
  ' Wynham Hotels
    ' Super 8 (El Dorado, Arkansas 71730)
    ' LaQuinta Inn (El Dorado, Arkansas 71730)

' --- The Purpose ---         
  ' To automate the launching process for SynXis PM - making it easier for employee's to continue working.
  ' To free up time normally used to go through the steps - thus increasing productivity since they will no longer have to do this manually.

' Instruction box before code is run.
MsgBox "Click [Okay] to load SynXis PM. It will automatically load itself; no action will be necessary. Please wait for the next popup.", 10, "SynXis Auto Loader"

' Create WScript.Shell object
Set x = CreateObject("WScript.Shell")

' Open SynXis link in a new Microsoft Edge tab named "Untitled"
x.Run "msedge https://whg.sabrehospitality.com/Loader/Loader.application?1=1&Lang=en-us"

' Wait until the new tab is open (check every 500 milliseconds)
Do While Not IsTabOpen("Untitled")
    WScript.Sleep 500
Loop

' Move cursor to "Open" and press it.
PressKeyAndWait x, "{LEFT}"
PressKeyAndWait x, "{ENTER}"

' Close the web browser (MS Edge)
x.Run "TaskKill /F /IM msedge.exe"

' Ensure the security warning popup is focused, select and press "Run".
FocusAndPressKey x, "Application Run - Security Warning", "%R"

' Enter a set of numbers and tab to a new line.
WaitUntilSuccessfulInput x, "12654", "{TAB}"

' Final popup to inform that the script has finished.
MsgBox "SynXis has been successfully loaded. You can now log in.", 0, "SynXis PM Auto Loader"

' SynXis should now be open, ready to use, and the browser closed. Employees can now log in.

Function PressKeyAndWait(shell, key)
    shell.SendKeys key
    WScript.Sleep 500
End Function

Function FocusAndPressKey(shell, windowTitle, key)
    shell.AppActivate windowTitle
    WScript.Sleep 500
    PressKeyAndWait shell, key
End Function

Sub WaitUntilSuccessfulInput(shell, inputText, keyAfterInput)
    ' Wait until the input is successfully entered (check every 500 milliseconds)
    Do While Not IsInputSuccessful(shell, inputText)
        WScript.Sleep 500
    Loop

    ' Press the key after successful input
    PressKeyAndWait shell, keyAfterInput
End Sub

Function IsInputSuccessful(shell, inputText)
    ' Add your conditions to check for successful input
    ' In this example, assuming successful input if the text is present in a specific location (adjust as needed)
    textFound = InStr(shell.SendKeys(inputText), "TextAfterInput")
    IsInputSuccessful = (textFound > 0)
End Function

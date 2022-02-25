Set WshShell = CreateObject("WScript.Shell")
Do
	WshShell.SendKeys ("{CAPSLOCK 2}")
	WScript.Sleep(60000)
Loop
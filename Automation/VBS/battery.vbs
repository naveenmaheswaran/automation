Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery",,48)
For Each objItem in colItems
    Sapi.speak "EstimatedChargeRemaining "
    Sapi.speak objItem.EstimatedChargeRemaining&"%"
    
Next
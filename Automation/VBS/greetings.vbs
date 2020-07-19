Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
 dim str
 if hour(time) < 12 then
 Sapi.speak "Good Morning Naveen "
 
 else
 if hour(time) > 12 then
 if hour(time) > 16 then
 Sapi.speak "Good evening Naveen "
 else
 Sapi.speak "Good afternoon Naveen "
 
 end if
 end if
 end if


Sapi.speak "Time is "
Sapi.speak time
Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
dim str
if hour(time) < 12 then
Sapi.speak "Good Morning sir"
else
if hour(time) > 12 then
if hour(time) > 16 then
Sapi.speak "Good evening sir "
else
Sapi.speak "Good afternoon sir "
end if
end if
end if


dim str1
set objTTS  = createobject("sapi.spvoice")
'str1 = "现在时间是" & Time
'objTTS.speak str1
while true
	if Minute(Time)/10=0 then
		str1 = "现在时间是" & Time
		objTTS.speak str1
	end if
wend
dim str1
set objTTS  = createobject("sapi.spvoice")
'str1 = "����ʱ����" & Time
'objTTS.speak str1
while true
	if Minute(Time)/10=0 then
		str1 = "����ʱ����" & Time
		objTTS.speak str1
	end if
wend
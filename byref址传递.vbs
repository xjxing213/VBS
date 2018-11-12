a=0
b=1

byRefDemo a,b
msgbox a & "--" & b

function byRefDemo(byval n1,byref n2)

	msgbox n1 & "--" & n2
	n1 = 10
	n2 = 20
	msgbox n1 & "--" & n2

end function
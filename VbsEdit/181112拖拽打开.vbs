dim arges,i,wsh

set arges = Wscript.Arguments
set wsh = createobject("Wscript.shell")
for i = 0 to arges.count - 1
	wsh.run "notepad " & arges.Item(i)
next




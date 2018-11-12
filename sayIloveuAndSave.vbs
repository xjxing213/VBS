dim wsh

set wsh = createobject("Wscript.Shell")

wsh.run "notepad"
wscript.sleep 100

wsh.sendkeys "+"
wscript.sleep 100

wsh.sendkeys "I"
wscript.sleep 1000

wsh.sendkeys " "
wscript.sleep 1000

wsh.sendkeys "L"
wscript.sleep 1000

wsh.sendkeys "o"
wscript.sleep 1000

wsh.sendkeys "v"
wscript.sleep 1000

wsh.sendkeys "e"
wscript.sleep 1000

wsh.sendkeys " "
wscript.sleep 1000

wsh.sendkeys "U"
wscript.sleep 1000

wsh.sendkeys "^(s)"
wscript.sleep 2000

wsh.sendkeys "vbsDemo"
wscript.sleep 1000

wsh.sendkeys "{ENTER}"
wscript.sleep 1000












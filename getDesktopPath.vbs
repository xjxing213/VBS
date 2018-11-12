dim wsh

set wsh = createobject("Wscript.Shell")

msgbox wsh.SpecialFolders("Desktop")
Option Explicit
Dim WshShell
Dim oExcel, oBook, oModule
Dim strRegKey, strCode, x, y
Set oExcel = CreateObject("Excel.Application") '���� Excel ����
set WshShell = CreateObject("wscript.Shell")
x = oExcel.Run("GetXCursorPos") '��ȡ��� X ����
y = oExcel.Run("GetYCursorPos") '��ȡ��� Y ����
WScript.Echo x, y
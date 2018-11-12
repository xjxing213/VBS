Option Explicit
Dim WshShell
Dim oExcel, oBook, oModule
Dim strRegKey, strCode, x, y
Set oExcel = CreateObject("Excel.Application") '创建 Excel 对象
set WshShell = CreateObject("wscript.Shell")
x = oExcel.Run("GetXCursorPos") '获取鼠标 X 坐标
y = oExcel.Run("GetYCursorPos") '获取鼠标 Y 坐标
WScript.Echo x, y
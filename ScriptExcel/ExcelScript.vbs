Option Explicit
Dim oExcelApp

Set oExcelApp = CreateObject("Excel.Application")
oExcelApp.Visible = True
oExcelApp.Workbooks.Add

Dim i
Randomize
For i = 1 To 10
oExcelApp.Cells(i, 1).Value = Int((30)*Rnd)
oExcelApp.Cells(i, 2).Value = Int((30)*Rnd)
oExcelApp.Cells(i, 3).Value = Int((30)*Rnd)
Next

oExcelApp.Range(oExcelApp.Cells(1, 1), oExcelApp.Cells(10, 3)).Select
oExcelApp.ActiveSheet.Shapes.AddChart.Select
oExcelApp.ActiveChart.ChartType = 4
oExcelApp.ActiveChart.ChartStyle = 26
oExcelApp.ActiveChart.SetElement(2)
oExcelApp.ActiveSheet.ChartObjects(1).Activate
oExcelApp.ActiveChart.ChartTitle.Text = "??????"
oExcelApp.ActiveChart.SeriesCollection(1).Name = "=""A"""
oExcelApp.ActiveChart.SeriesCollection(2).Name = "=""B"""
oExcelApp.ActiveChart.SeriesCollection(3).Name = "=""C"""

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

oExcelApp.ActiveWorkbook.SaveAs(FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + "Table.xlsx")

WScript.Echo "?????????? ?? ????????? ? ???????: " + FSO.GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "?????? ???? ? ???????: " + FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + WScript.ScriptName

WScript.Echo "?????????? ???????:"
Dim f, str
Set f = FSO.OpenTextFile(WScript.ScriptName, 1)
Do While Not F.AtEndOfStream
	str = f.ReadLine
	WScript.Echo str
Loop
f.Close
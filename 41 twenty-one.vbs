Option Explicit

sExcelSample()

Sub sExcelSample()

	Dim objExcel
	Dim objRange
	
	Set objExcel = CreateObject("Excel.Application")
	
	objExcel.Visible = True
	
	objExcel.Workbooks.Add
	
	Set objRange = objExcel.Worksheets(1).Range("A1:G21")
	
	Set objRange = objExcel.Worksheets(1).Range("A1:G1")
	
	With objRange.Interior
		.ColorIndex = 35
    End With
    
    objExcel.Worksheets(1).Rows("1:1").RowHeight = 28.5
	WScript.Sleep(5000)

	objExcel.DisplayAlerts = False
	objExcel.Workbooks(1).SaveAs "C:\Users\bit-surf\Documents\MAKIYAMA\VBScript_Tutorial\work\testbook1.xlsx"
	objExcel.Workbooks(1).SaveAs "C:\Users\bit-surf\Documents\MAKIYAMA\VBScript_Tutorial\work\testbook2.xlsx"
	
	objExcel.Quit

	Set objExcel = Nothing

End Sub

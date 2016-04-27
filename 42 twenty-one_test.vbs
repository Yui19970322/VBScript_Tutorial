Option Explicit

sExcelSample()

Sub sExcelSample()

	Dim objExcel
	Dim objRange
	Dim i
	
	
	Set objExcel = CreateObject("Excel.Application")
	
	objExcel.Visible = True
	
	objExcel.Workbooks.Add
	
	Set objRange = objExcel.Worksheets(1).Range("A1:G21")
	
	Set objRange = objExcel.Worksheets(1).Range("A1:G1")
	
	With objRange.Interior
		.ColorIndex = 35
    End With
    
    objExcel.Worksheets(1).Rows("1:1").RowHeight = 28.5
	WScript.Sleep(1000)

	For i = 1 To 10
		WScript.echo i & "objExcel.Workbooks(1).SaveAs"
		objExcel.DisplayAlerts = False
		objExcel.Workbooks(1).SaveAs "C:\Users\bit-surf\Documents\MAKIYAMA\VBScript_Tutorial\work\testbook" &  i & ".xlsx"
	Next
	objExcel.Quit
	Set objExcel = Nothing
End Sub

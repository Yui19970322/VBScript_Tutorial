Option Explicit

Dim objExcel

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

WScript.Sleep(5000)

objExcel.DisplayAlerts = False
objExcel.Quit

Set objExcel = Nothing


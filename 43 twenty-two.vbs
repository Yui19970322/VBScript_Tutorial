Option Explicit

sWordSample()
Sub sWordSample()

	Dim objWord
	Dim objSelection
	
	
	Set objWord = CreateObject("Word.Application")
	
	
	objWord.Visible = True
	
	objWord.Documents.Add
	
    objWord.ActiveWindow.Selection.TypeText "Ç®ÇÕÇÊÇ§Ç≤Ç¥Ç¢Ç‹Ç∑"
    objWord.ActiveWindow.Selection.WholeStory
    With objWord.Selection
        .Font.Size = 20
	    .Font.Name = "MSÉSÉVÉbÉN"
    End With

    

WScript.Sleep(5000)

objWord.DisplayAlerts = False



objWord.Documents(1).SaveAs "C:\Users\bit-surf\Documents\MAKIYAMA\VBScript_Tutorial\work\testdocument.doc"



objWord.Quit 
Set objWord = Nothing

End Sub

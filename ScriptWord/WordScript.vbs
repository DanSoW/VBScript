Option Explicit

Dim oWord, oDoc, oSelectionPlace

Set oWord = CreateObject("Word.Application")
oWord.Visible = True

Set oDoc = oWord.Documents.Add()
Set oSelectionPlace = oWord.Selection

oSelectionPlace.TypeParagraph() 
oSelectionPlace.Paragraphs.Alignment = 1
oSelectionPlace.Font.Name = "ComicSans" 
oSelectionPlace.Font.Size = "18" 
oSelectionPlace.Font.Bold = True 
oSelectionPlace.Font.Color = RGB(0, 255, 0) 
oSelectionPlace.TypeText "Первый лист - ФИО"
 
oSelectionPlace.TypeParagraph() 
oSelectionPlace.TypeText "Солопов Даниил Дмитриевич"
oSelectionPlace.EndKey(6)
oSelectionPlace.InsertBreak

oSelectionPlace.TypeText "Второй лист - Номер варианта"
oSelectionPlace.TypeParagraph() 
oSelectionPlace.TypeText "Вариант №24"
oSelectionPlace.InsertBreak

oSelectionPlace.TypeText "Третий лист - Номер лабораторной работы"
oSelectionPlace.TypeParagraph() 
oSelectionPlace.TypeText "Лабораторная работа №2"

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
oWord.ActiveDocument.SaveAs(FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + "Document.docx")

WScript.Echo "Директория со скриптами и файлами: " + FSO.GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "Полный путь к скрипту: " + FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + WScript.ScriptName

WScript.Echo "Содержимое скрипта:"
Dim f, str
Set f = FSO.OpenTextFile(WScript.ScriptName, 1)
Do While Not F.AtEndOfStream
	str = f.ReadLine
	WScript.Echo str
Loop
f.Close
Option Explicit

Dim objPPT, objPresentation, objSlide, objTitle, objShapes
Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
'objPresentation.ApplyTemplate("C:\Program Files\Microsoft Office\Templates\1033\Pitchbook.potx")

Set objSlide = objPresentation.Slides.Add(1, 2)
Set objShapes = objSlide.Shapes
Set objTitle = objShapes.Item(1)
objTitle.TextFrame.TextRange.Text = "������ ����� - ���"
Set objTitle = objShapes.Item(2)
objTitle.TextFrame.TextRange.Text = "������� ������ ����������"

Set objSlide = objPresentation.Slides.Add(2, 2)
Set objShapes = objSlide.Shapes
Set objTitle = objShapes.Item(1)
objTitle.TextFrame.TextRange.Text = "������ ����� - ����� ��������"
Set objTitle = objShapes.Item(2)
objTitle.TextFrame.TextRange.Text = "������� �24"

Set objSlide = objPresentation.Slides.Add(3, 2)
Set objShapes = objSlide.Shapes
Set objTitle = objShapes.Item(1)
objTitle.TextFrame.TextRange.Text = "������ ����� - ����� ������������ ������"
Set objTitle = objShapes.Item(2)
objTitle.TextFrame.TextRange.Text = "������������ ������ �2"

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

objPresentation.SaveAs(FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + "Presentation.ppt")

WScript.Echo "���������� �� ��������� � �������: " + FSO.GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "������ ���� � �������: " + FSO.GetParentFolderName(WScript.ScriptFullName) + "\" + WScript.ScriptName

WScript.Echo "���������� �������:"
Dim f, str
Set f = FSO.OpenTextFile(WScript.ScriptName, 1)
Do While Not F.AtEndOfStream
	str = f.ReadLine
	WScript.Echo str
Loop
f.Close
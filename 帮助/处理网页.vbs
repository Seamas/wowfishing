Set fso = createobject("scripting.filesystemobject")
curdir = fso.getparentfoldername(wscript.scriptfullname)

Const base = "<base target=""_blank"">"
Const head = "<head>"
msg = ""

Set objfolder = fso.getfolder(curdir)

For Each objfile In objfolder.files
	If LCase(fso.getextensionname(objfile.name)) = "htm" Then
		processfile objfile.path
	End If
Next

msgbox msg & "ȫ���������!"

Sub ProcessFile(ByVal filepath)
	Set stream = fso.opentextfile(filepath,1,False)
	content = stream.readall
	stream.close

	If InStr(content,base) > 0 Then
		msg = msg & filepath & " �Ѵ����!" & vbcrlf
		Exit Sub
	End If

	content = Replace(content,head,head & vbcrlf & base)

	Set stream = fso.opentextfile(filepath,2,False)
	stream.write content
	stream.close
	msg = msg & filepath & " �������!" & vbcrlf

End Sub

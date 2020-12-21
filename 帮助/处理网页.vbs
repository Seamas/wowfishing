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

msgbox msg & "全部处理完成!"

Sub ProcessFile(ByVal filepath)
	Set stream = fso.opentextfile(filepath,1,False)
	content = stream.readall
	stream.close

	If InStr(content,base) > 0 Then
		msg = msg & filepath & " 已处理过!" & vbcrlf
		Exit Sub
	End If

	content = Replace(content,head,head & vbcrlf & base)

	Set stream = fso.opentextfile(filepath,2,False)
	stream.write content
	stream.close
	msg = msg & filepath & " 处理完成!" & vbcrlf

End Sub

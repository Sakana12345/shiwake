a = InputBox ("���O����͂��Ă�������","�`���b�g(vba)")
If IsEmpty(a) then
MsgBox "�{���ɕ��܂����H"
WScript.Quit

end if


Dim objFileSys
Dim strScriptPath
Dim strCreateFile

Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,a+("���������܂���"))

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing

 

do

x = InputBox ("�������������Ƃ���͂��Ă�������","�`���b�g")
If IsEmpty(x) then
Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,a+("���ގ����܂���"))

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing
MsgBox "�I�����܂�"
WScript.Quit

end if

 


Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,x+("�@by")+a)

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing

loop
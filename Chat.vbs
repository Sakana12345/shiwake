a = InputBox ("名前を入力してください","チャット(vba)")
If IsEmpty(a) then
MsgBox "本当に閉じますか？"
WScript.Quit

end if


Dim objFileSys
Dim strScriptPath
Dim strCreateFile

Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,a+("が入室しました"))

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing

 

do

x = InputBox ("発言したいことを入力してください","チャット")
If IsEmpty(x) then
Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,a+("が退室しました"))

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing
MsgBox "終了します"
WScript.Quit

end if

 


Set objFileSys = CreateObject("Scripting.FileSystemObject")

strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

 

On Error Resume Next

 

strCreateFile = objFileSys.BuildPath(strScriptPath,x+("　by")+a)

objFileSys.CreateTextFile strCreateFile

 

Set objFileSys = Nothing

loop
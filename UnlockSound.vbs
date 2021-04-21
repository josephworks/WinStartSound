Set WshShell = CreateObject("WScript.shell")
'cycle 30 times
for j = 0 to 1
'crank it back up 
for i = 0 to 50
WshShell.SendKeys(chr(175))
next

next

Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Mogus.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
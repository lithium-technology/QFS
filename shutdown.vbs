Dim IntCounter 
Dim objWshShl : set objWshShl = WScript.CreateObject("wscript.shell")
Dim objVoice : set objVoice = WScript.CreateObject("sapi.spvoice")
ShutdownWarning()
TimedMessageBox()
ShutdownComputer()
Function ShutdownWarning
    objVoice.Speak "Computer shut down from 10 sec."
    WScript.Sleep 5000
End Function
Function TimedMessageBox
    For IntCounter = 5 To 1 Step -1
        objWshShl.Popup "Computer will shutdown in" _
        & IntCounter & " seconds",1,"Computer Shutdown", 0+48
    Next
End Function
objWshShl.Popup "Computer will shutdown in " _
        & IntCounter & " seconds",1,"Computer Shutdown", 0+48
Function ShutdownComputer
    objWshShl.Run "Shutdown /s /f /t 0",0
End Function

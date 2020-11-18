Option Explicit

' ------------------------------------------------------------------------- 
' Autor: Isac Canedo
' Criado: 28/08/2003
' -------------------------------------------------------------------------

host="127.0.0.1"

do
 WScript.Sleep (1000*60)
loop While (ping(host))

WScript.Run "resume.txt"
WScript.Quit

Function ping(strComputer)
Dim ObjShell, objScriptExec, strComputer, strPingResults, success
 Set objShell = CreateObject("WScript.Shell")
 Set objScriptExec = objShell.Exec( "ping -n 2 -w 1000 " & strComputer)
  strPingResults = LCase(objScriptExec.StdOut.ReadAll)
  If InStr(strPingResults, "reply from") Then
     If InStr(strPingResults, "destination net unreachable") Then
         success=False
     Else
         success=True
     End If 
 Else
    success=False
 End If
ping=success
End Function
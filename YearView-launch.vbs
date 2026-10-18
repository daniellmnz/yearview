' YearView hidden launcher.
' wscript.exe is a windowless (GUI-subsystem) host, so launching PowerShell from
' here with window style 0 never paints a console - unlike a shortcut pointing at
' powershell.exe, which creates its console box before -WindowStyle Hidden applies
' (and which Windows Terminal may ignore entirely). The script's own hidden
' respawn still runs as a second line of defence on hosts that ignore style 0.
' Any arguments given to this script are passed through (default: -Serve -App).
Option Explicit
Dim sh, fso, dir, args, a, cmd
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
args = " -Serve -App"
If WScript.Arguments.Count > 0 Then
    args = ""
    For Each a In WScript.Arguments
        If InStr(a, " ") > 0 Then
            args = args & " """ & a & """"
        Else
            args = args & " " & a
        End If
    Next
End If
cmd = """" & sh.ExpandEnvironmentStrings("%SystemRoot%") & _
      "\System32\WindowsPowerShell\v1.0\powershell.exe""" & _
      " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & dir & _
      "\New-OutlookYearView.ps1""" & args
sh.Run cmd, 0, False

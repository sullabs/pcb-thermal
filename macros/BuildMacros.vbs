' Builds PCBThermal.swp from the .bas files next to it, through SolidWorks (started when it is not running).
' Run it with BuildMacros.cmd. Result: %TEMP%\PCBThermal_build.log, exit code 0 = OK.
' BuildMacros.main imports; the VBA editor compiles and saves after the macro has ended, so this waits for the new
' .swp and then runs BuildMacros.Check.
Option Explicit
Dim fso, sh, sw, here, swp, logF, ok, t, txt, stamp
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
here = fso.GetParentFolderName(WScript.ScriptFullName)
swp = here & "\PCBThermal.swp"
logF = sh.ExpandEnvironmentStrings("%TEMP%") & "\PCBThermal_build.log"

On Error Resume Next
Set sw = GetObject(, "SldWorks.Application")
On Error GoTo 0
If IsEmpty(sw) Then
    WScript.Echo "Starting SolidWorks..."
    Set sw = CreateObject("SldWorks.Application")
    sw.Visible = True
    t = Timer
    Do While Not sw.StartupProcessCompleted And Timer - t < 300
        WScript.Sleep 1000
    Loop
End If

If fso.FileExists(logF) Then fso.DeleteFile logF
stamp = fso.GetFile(swp).DateLastModified
ok = sw.RunMacro(swp, "BuildMacros", "main")
If Not fso.FileExists(logF) Then
    ' First time: PCBThermal.swp has no BuildMacros module yet. Update Macros imports it.
    WScript.Echo "PCBThermal.swp has no BuildMacros yet. Running Update Macros: click OK in SolidWorks when it is done."
    sw.RunMacro swp, "UpdateMacros", "main"
    stamp = fso.GetFile(swp).DateLastModified
    ok = sw.RunMacro(swp, "BuildMacros", "main")
End If
If Not fso.FileExists(logF) Then
    WScript.Echo "FAIL BuildMacros did not run (RunMacro returned " & ok & ")."
    WScript.Quit 2
End If

t = Timer
Do While fso.GetFile(swp).DateLastModified = stamp And Timer - t < 60
    WScript.Sleep 1000
Loop
If fso.GetFile(swp).DateLastModified = stamp Then
    fso.OpenTextFile(logF, 8).WriteLine "FAIL not saved: look for a compile error in the VBA editor (Alt+F11)"
Else
    fso.OpenTextFile(logF, 8).WriteLine "saved"
    sw.RunMacro swp, "BuildMacros", "Check"
End If
txt = fso.OpenTextFile(logF, 1).ReadAll
WScript.Echo txt
If InStr(txt, "FAIL") = 0 And InStr(txt, "OK") > 0 Then WScript.Quit 0 Else WScript.Quit 1

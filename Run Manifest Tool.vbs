' ============================================
' Multi-Carrier Manifest Tool Launcher
' ============================================
' Shows startup status, auto-updates from Git,
' and launches the GUI even if update fails.
' ============================================

Option Explicit

Dim FSO, WshShell
Dim strScriptPath, strLogFile
Dim statusFile, htaFile
Dim loaderExec
Dim updateTimeoutSeconds

updateTimeoutSeconds = 25

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Set working directory to script location
strScriptPath = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strScriptPath
strLogFile = strScriptPath & "\update_errors.log"

' Temp files for startup status window
statusFile = FSO.GetSpecialFolder(2) & "\manifest_tool_startup_status.txt"
htaFile = FSO.GetSpecialFolder(2) & "\manifest_tool_startup.hta"

Call WriteLoaderHta(htaFile, statusFile)
Set loaderExec = WshShell.Exec("mshta """ & htaFile & """")

Call UpdateStatus("Starting Multi-Carrier Manifest Tool" & vbCrLf & _
                  "Preparing startup checks...")
WScript.Sleep 250

' Run git pull silently and capture result
Call UpdateStatus("Checking for updates..." & vbCrLf & _
                  "Running git pull (up to " & updateTimeoutSeconds & " seconds).")

Dim objExec, startTime, timedOut
Set objExec = WshShell.Exec("cmd /c git pull 2>&1")
startTime = Now
timedOut = False

Do While objExec.Status = 0
    WScript.Sleep 200
    If DateDiff("s", startTime, Now) >= updateTimeoutSeconds Then
        timedOut = True
        Exit Do
    End If
Loop

If timedOut Then
    Call AppendUpdateLog( _
        "Update timed out after " & updateTimeoutSeconds & " seconds." & vbCrLf & _
        "Launching app without waiting for git pull to finish." _
    )

    On Error Resume Next
    WshShell.Run "cmd /c taskkill /F /PID " & objExec.ProcessID & " >nul 2>&1", 0, True
    On Error GoTo 0
ElseIf objExec.ExitCode <> 0 Then
    Dim strOutput, strError
    strOutput = ""
    strError = ""

    On Error Resume Next
    strOutput = objExec.StdOut.ReadAll
    strError = objExec.StdErr.ReadAll
    On Error GoTo 0

    Call AppendUpdateLog( _
        "Update failed." & vbCrLf & _
        "Exit code: " & objExec.ExitCode & vbCrLf & _
        "Output: " & strOutput & vbCrLf & _
        "Error: " & strError _
    )
End If

Call UpdateStatus("Starting application..." & vbCrLf & _
                  "Loading interface and startup checks.")

' Launch the GUI regardless of update result
WshShell.Run "pythonw gui.py", 0, False

Call CloseLoader()


Sub UpdateStatus(message)
    On Error Resume Next
    Dim f
    Set f = FSO.OpenTextFile(statusFile, 2, True) ' 2 = overwrite
    f.Write message
    f.Close
    On Error GoTo 0
End Sub


Sub CloseLoader()
    On Error Resume Next
    Call UpdateStatus("__CLOSE__")
    WScript.Sleep 350

    If Not loaderExec Is Nothing Then
        If loaderExec.Status = 0 Then
            WshShell.Run "cmd /c taskkill /F /PID " & loaderExec.ProcessID & " >nul 2>&1", 0, True
        End If
    End If

    If FSO.FileExists(statusFile) Then
        FSO.DeleteFile statusFile, True
    End If
    If FSO.FileExists(htaFile) Then
        FSO.DeleteFile htaFile, True
    End If
    On Error GoTo 0
End Sub


Sub AppendUpdateLog(message)
    On Error Resume Next
    Dim objLog
    Set objLog = FSO.OpenTextFile(strLogFile, 8, True) ' 8 = append mode
    objLog.WriteLine "============================================"
    objLog.WriteLine "Update check: " & Now()
    objLog.WriteLine message
    objLog.WriteLine ""
    objLog.Close
    On Error GoTo 0
End Sub


Sub WriteLoaderHta(targetPath, statusPath)
    Dim f
    Set f = FSO.OpenTextFile(targetPath, 2, True)

    f.WriteLine "<html>"
    f.WriteLine "<head>"
    f.WriteLine "<title>Starting Manifest Tool</title>"
    f.WriteLine "<HTA:APPLICATION ID=""ManifestLoader"""
    f.WriteLine "  APPLICATIONNAME=""ManifestLoader"""
    f.WriteLine "  BORDER=""thin"""
    f.WriteLine "  CAPTION=""yes"""
    f.WriteLine "  SHOWINTASKBAR=""yes"""
    f.WriteLine "  SINGLEINSTANCE=""yes"""
    f.WriteLine "  WINDOWSTATE=""normal"">"
    f.WriteLine "<style>"
    f.WriteLine "body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; margin: 0; padding: 18px; }"
    f.WriteLine "h2 { margin: 0 0 10px 0; color: #1f2937; font-size: 20px; }"
    f.WriteLine ".panel { background: #ffffff; border: 1px solid #d9dee5; border-radius: 8px; padding: 14px; }"
    f.WriteLine "pre { margin: 0; white-space: pre-wrap; font-size: 13px; color: #111827; }"
    f.WriteLine ".hint { margin-top: 10px; color: #6b7280; font-size: 12px; }"
    f.WriteLine "</style>"
    f.WriteLine "<script language=""VBScript"">"
    f.WriteLine "Option Explicit"
    f.WriteLine "Dim loaderFso"
    f.WriteLine "Dim statusPath"
    f.WriteLine "Sub Window_OnLoad"
    f.WriteLine "  Set loaderFso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "  statusPath = """ & Replace(statusPath, """", """""") & """"
    f.WriteLine "  window.resizeTo 560, 250"
    f.WriteLine "  window.moveTo (screen.availWidth - 560) \ 2, (screen.availHeight - 250) \ 2"
    f.WriteLine "  window.setInterval ""RefreshStatus"", 300"
    f.WriteLine "  RefreshStatus"
    f.WriteLine "End Sub"
    f.WriteLine "Sub RefreshStatus"
    f.WriteLine "  On Error Resume Next"
    f.WriteLine "  If loaderFso.FileExists(statusPath) Then"
    f.WriteLine "    Dim statusFileHandle, txt"
    f.WriteLine "    Set statusFileHandle = loaderFso.OpenTextFile(statusPath, 1, False)"
    f.WriteLine "    txt = statusFileHandle.ReadAll"
    f.WriteLine "    statusFileHandle.Close"
    f.WriteLine "    document.getElementById(""status"").innerText = txt"
    f.WriteLine "    If InStr(txt, ""__CLOSE__"") > 0 Then window.close"
    f.WriteLine "  End If"
    f.WriteLine "End Sub"
    f.WriteLine "</script>"
    f.WriteLine "</head>"
    f.WriteLine "<body>"
    f.WriteLine "  <h2>Multi-Carrier Manifest Tool</h2>"
    f.WriteLine "  <div class=""panel"">"
    f.WriteLine "    <pre id=""status"">Starting...</pre>"
    f.WriteLine "  </div>"
    f.WriteLine "  <div class=""hint"">Tip: startup can be slower if VPN/network folders are slow.</div>"
    f.WriteLine "</body>"
    f.WriteLine "</html>"

    f.Close
End Sub

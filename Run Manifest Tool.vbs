' ============================================
' Multi-Carrier Manifest Tool Launcher
' ============================================
' Auto-updates from Git before launching
' Logs update failures to update_errors.log
' ============================================

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Set working directory to script location
strScriptPath = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strScriptPath

' Run git pull silently and capture result
strLogFile = strScriptPath & "\update_errors.log"
strCommand = "cmd /c git pull 2>&1"

Set objExec = WshShell.Exec(strCommand)

' Wait for git pull to complete
Do While objExec.Status = 0
    WScript.Sleep 100
Loop

' Check if git pull failed (non-zero exit code)
If objExec.ExitCode <> 0 Then
    ' Read the error output
    strOutput = objExec.StdOut.ReadAll
    strError = objExec.StdErr.ReadAll

    ' Log the failure with timestamp
    Set objLog = FSO.OpenTextFile(strLogFile, 8, True) ' 8 = append mode
    objLog.WriteLine "============================================"
    objLog.WriteLine "Update failed: " & Now()
    objLog.WriteLine "Exit code: " & objExec.ExitCode
    If Len(strOutput) > 0 Then
        objLog.WriteLine "Output: " & strOutput
    End If
    If Len(strError) > 0 Then
        objLog.WriteLine "Error: " & strError
    End If
    objLog.WriteLine ""
    objLog.Close
End If

' Launch the GUI regardless of update result
WshShell.Run "pythonw gui.py", 0, False

' Outlook COM Add-in - Simple VBScript COM Add-in
' This creates a proper COM add-in that Outlook can load

Option Explicit

' Global variables
Dim outlookApp
Dim outlookNamespace

' Main connection function
Function OnConnection(application, connectMode, addInInst, custom)
    On Error Resume Next
    
    Set outlookApp = application
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    
    ' Write success log
    WriteLog "AI Email Manager add-in connected to Outlook successfully!"
    WriteLog "Connect Mode: " & connectMode
    
    OnConnection = True
End Function

' Disconnection function
Function OnDisconnection(removeMode, custom)
    On Error Resume Next
    
    WriteLog "AI Email Manager add-in disconnected from Outlook"
    
    OnDisconnection = True
End Function

' Startup complete function
Function OnStartupComplete(custom)
    On Error Resume Next
    
    WriteLog "AI Email Manager add-in ready for use!"
    
    OnStartupComplete = True
End Function

' Shutdown function
Function OnBeginShutdown(custom)
    On Error Resume Next
    
    WriteLog "AI Email Manager add-in shutting down..."
    
    OnBeginShutdown = True
End Function

' Helper function to write logs
Sub WriteLog(message)
    On Error Resume Next
    
    Dim fso, logFile, logPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the script directory
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\outlook_addin_success.log"
    
    ' Open log file for appending
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8 = ForAppending
    Else
        Set logFile = fso.CreateTextFile(logPath, True)
    End If
    
    logFile.WriteLine Now() & " - " & message
    logFile.Close
End Sub





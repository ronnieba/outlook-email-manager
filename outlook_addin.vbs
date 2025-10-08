' Outlook VBA Add-in - Simple working add-in
' This is a VBScript that can be used as a COM add-in

Option Explicit

Dim outlookApp
Dim namespace

' Function to connect to Outlook
Function OnConnection(application, connectMode, addInInst, custom)
    On Error Resume Next
    
    Set outlookApp = application
    Set namespace = outlookApp.GetNamespace("MAPI")
    
    ' Write to log file
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.CreateTextFile("outlook_addin_success.log", True)
    logFile.WriteLine "AI Email Manager add-in connected to Outlook successfully!"
    logFile.WriteLine "Connect Mode: " & connectMode
    logFile.Close
    
    OnConnection = True
End Function

' Function to disconnect from Outlook
Function OnDisconnection(removeMode, custom)
    On Error Resume Next
    
    ' Write to log file
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile("outlook_addin_success.log", 8, True)
    logFile.WriteLine "AI Email Manager add-in disconnected from Outlook"
    logFile.Close
    
    OnDisconnection = True
End Function

' Function to start add-in
Function OnStartupComplete(custom)
    On Error Resume Next
    
    ' Write to log file
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile("outlook_addin_success.log", 8, True)
    logFile.WriteLine "AI Email Manager add-in ready for use!"
    logFile.Close
    
    OnStartupComplete = True
End Function

' Function to end add-in
Function OnBeginShutdown(custom)
    On Error Resume Next
    
    ' Write to log file
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile("outlook_addin_success.log", 8, True)
    logFile.WriteLine "AI Email Manager add-in shutting down..."
    logFile.Close
    
    OnBeginShutdown = True
End Function
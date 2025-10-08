# Outlook COM Add-in - PowerShell COM Add-in
# This creates a proper COM add-in using PowerShell

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("12345678-1234-1234-1234-123456789012")]
public class OutlookAddin
{
    private dynamic outlookApp;
    private dynamic namespace;
    
    public bool OnConnection(dynamic application, int connectMode, dynamic addInInst, dynamic custom)
    {
        try
        {
            outlookApp = application;
            namespace = outlookApp.GetNamespace("MAPI");
            
            // Write to log file
            System.IO.File.WriteAllText("outlook_addin_success.log", 
                "AI Email Manager add-in connected to Outlook successfully!\n" +
                "Connect Mode: " + connectMode + "\n");
            
            return true;
        }
        catch (Exception e)
        {
            System.IO.File.WriteAllText("outlook_addin_error.log", 
                "Error connecting to Outlook: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnDisconnection(int removeMode, dynamic custom)
    {
        try
        {
            System.IO.File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in disconnected from Outlook\n");
            return true;
        }
        catch (Exception e)
        {
            System.IO.File.AppendAllText("outlook_addin_error.log", 
                "Error disconnecting from Outlook: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnStartupComplete(dynamic custom)
    {
        try
        {
            System.IO.File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in ready for use!\n");
            return true;
        }
        catch (Exception e)
        {
            System.IO.File.AppendAllText("outlook_addin_error.log", 
                "Error in startup: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnBeginShutdown(dynamic custom)
    {
        try
        {
            System.IO.File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in shutting down...\n");
            return true;
        }
        catch (Exception e)
        {
            System.IO.File.AppendAllText("outlook_addin_error.log", 
                "Error in shutdown: " + e.Message + "\n");
            return false;
        }
    }
}
"@

# Register the COM add-in
$regasm = Get-Command regasm -ErrorAction SilentlyContinue
if ($regasm) {
    # Compile and register
    Add-Type -Path "outlook_addin.cs" -OutputAssembly "outlook_addin.dll"
    regasm "outlook_addin.dll" /codebase
    Write-Host "COM add-in registered successfully!"
} else {
    Write-Host "regasm not found. Please install .NET Framework SDK."
}





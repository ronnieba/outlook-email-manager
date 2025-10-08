using System;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop.Outlook;

[ComVisible(true)]
[Guid("12345678-1234-1234-1234-123456789012")]
[ClassInterface(ClassInterfaceType.None)]
public class OutlookAddin : IDTExtensibility2
{
    private Application outlookApp;
    private NameSpace outlookNamespace;
    
    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            outlookApp = (Application)Application;
            outlookNamespace = outlookApp.GetNamespace("MAPI");
            
            // Write to log file
            File.WriteAllText("outlook_addin_success.log", 
                "AI Email Manager add-in connected to Outlook successfully!\n" +
                "Connect Mode: " + ConnectMode + "\n");
        }
        catch (Exception e)
        {
            File.WriteAllText("outlook_addin_error.log", 
                "Error connecting to Outlook: " + e.Message + "\n");
        }
    }
    
    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in disconnected from Outlook\n");
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error disconnecting from Outlook: " + e.Message + "\n");
        }
    }
    
    public void OnAddInsUpdate(ref Array custom)
    {
        // Not implemented
    }
    
    public void OnStartupComplete(ref Array custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in ready for use!\n");
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error in startup: " + e.Message + "\n");
        }
    }
    
    public void OnBeginShutdown(ref Array custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in shutting down...\n");
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error in shutdown: " + e.Message + "\n");
        }
    }
}





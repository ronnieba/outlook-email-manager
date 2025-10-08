using System;
using System.Runtime.InteropServices;
using System.IO;

[ComVisible(true)]
[Guid("12345678-1234-1234-1234-123456789012")]
public class OutlookAddin
{
    private object outlookApp;
    private object outlookNamespace;
    
    public bool OnConnection(object application, int connectMode, object addInInst, object custom)
    {
        try
        {
            outlookApp = application;
            outlookNamespace = ((dynamic)outlookApp).GetNamespace("MAPI");
            
            // Write to log file
            File.WriteAllText("outlook_addin_success.log", 
                "AI Email Manager add-in connected to Outlook successfully!\n" +
                "Connect Mode: " + connectMode + "\n");
            
            return true;
        }
        catch (Exception e)
        {
            File.WriteAllText("outlook_addin_error.log", 
                "Error connecting to Outlook: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnDisconnection(int removeMode, object custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in disconnected from Outlook\n");
            return true;
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error disconnecting from Outlook: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnStartupComplete(object custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in ready for use!\n");
            return true;
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error in startup: " + e.Message + "\n");
            return false;
        }
    }
    
    public bool OnBeginShutdown(object custom)
    {
        try
        {
            File.AppendAllText("outlook_addin_success.log", 
                "AI Email Manager add-in shutting down...\n");
            return true;
        }
        catch (Exception e)
        {
            File.AppendAllText("outlook_addin_error.log", 
                "Error in shutdown: " + e.Message + "\n");
            return false;
        }
    }
}





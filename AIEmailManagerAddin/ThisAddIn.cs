using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace AIEmailManagerAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // רישום השדות המותאמים אישית בכל התיקיות
                RegisterCustomFields();
            }
            catch
            {
                // אם יש שגיאה, המשך בלי לקרוס
            }
        }

        private void RegisterCustomFields()
        {
            try
            {
                var inbox = Application.GetNamespace("MAPI").GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                
                // רשום AISCORE
                if (inbox.UserDefinedProperties.Find("AISCORE") == null)
                {
                    inbox.UserDefinedProperties.Add("AISCORE", 
                        Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText);
                }

                // רשום AI Analysis
                if (inbox.UserDefinedProperties.Find("AI Analysis") == null)
                {
                    inbox.UserDefinedProperties.Add("AI Analysis", 
                        Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText);
                }

                // רשום AICategory
                if (inbox.UserDefinedProperties.Find("AICategory") == null)
                {
                    inbox.UserDefinedProperties.Add("AICategory", 
                        Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText);
                }
            }
            catch
            {
                // אם השדות כבר קיימים או יש שגיאה, המשך
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

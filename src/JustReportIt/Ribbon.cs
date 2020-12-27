using JustReportIt.Properties;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JustReportIt
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        private const string LOOKUP_URL = "https://api.justreport.it/lookup/";
        private const string API_KEY = "DeVKFVPs3T40XYrUeQl9z5adMwopbYnY8jaAbecw";

        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public void OnTextButton(Office.IRibbonControl control)
        {
            Outlook.MailItem item = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            try
            {
                CreateMailItem(item);
            }
            catch
            {
                MessageBox.Show("WHOIS lookup has failed", "Error");
            }

        }

        private void CreateMailItem(Outlook.MailItem selectedItem)
        {
            Outlook.PropertyAccessor olPA = selectedItem.PropertyAccessor;
            string header = olPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS);
            string spamDomain = selectedItem.SenderEmailAddress.Split('@')[1];
            POCO response = JsonConvert.DeserializeObject<POCO>(GetAbuseEmail(spamDomain));
            Outlook.MailItem mailItem = (Outlook.MailItem)
                Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "Spam Abuse from: " + spamDomain;
            mailItem.To = response.data["email"];
            mailItem.Body = GetBody(selectedItem.Body, header, spamDomain);
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        private String GetBody(string body, string header, string spamDomain)
        {
            return "To whom it may concern, \n\n" +
                "I am writing to you today to report the following domain: " + spamDomain + "\n\n" +
                "This domain is sending me unsolicited spam emails. \n\n" +
                "Please take appropriate measures to avoid future abuse from this user. \n\n" +
                "You will find the raw spam email below: \n\n" + header + body;
        }

        public string GetAbuseEmail(string abuseDomain)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(LOOKUP_URL + abuseDomain);
            request.Headers.Add("x-api-key", API_KEY);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public Bitmap imageSuper_GetImage(IRibbonControl control)
        {
            return Resources.spam;
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("JustReportIt.Ribbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}

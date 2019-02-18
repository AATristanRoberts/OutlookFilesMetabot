using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookFiles
{
    public class OutlookFiles
    {
        static private Microsoft.Office.Interop.Outlook.Application outlook;
        static private Dictionary<String, MailItem> openMessages = new Dictionary<string, MailItem>();

        public void OpenOutlook()
        {
            outlook = new Microsoft.Office.Interop.Outlook.Application();
        }

        public void CloseOutlook()
        {
            outlook.Quit();
            outlook = null;
        }

        public string OpenFile(string path)
        {
            MailItem msg = outlook.Session.OpenSharedItem(path);
            string key = msg.EntryID == null ? path : msg.EntryID;
            openMessages.Add(key, msg);
            return key;
        }

        public void CloseFile(string entryID)
        {
            if (openMessages.ContainsKey(entryID))
            {
                openMessages[entryID].Close(OlInspectorClose.olDiscard);
            }
        }

        public string GetSubject(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].Subject;
        }

        public string GetSenderName(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].SenderName;
        }

        public string GetSenderEmail(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].SenderEmailAddress;
        }

        public string GetSentDate(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].ReceivedTime.ToString();
        }

        public string GetTo(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].To;
        }

        public string GetCc(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].CC;
        }

        public string GetBcc(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].BCC;
        }

        public int GetAttachmentCount(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return -1;
            }

            return openMessages[entryID].Attachments.Count;
        }

        public string GetAttachmentName(string entryID, int n)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].Attachments[n].FileName;
        }

        public int GetAttachmentSize(string entryID, int n)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return -1;
            }

            return openMessages[entryID].Attachments[n].Size;
        }

        public string GetAttachmentType(string entryID, int n)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].Attachments[n].Type.ToString();
        }

        public void SaveAttachmentAs(string entryID, int n, string savePath)
        {
            if (openMessages.ContainsKey(entryID))
            {
                openMessages[entryID].Attachments[n].SaveAsFile(savePath);
            }
        }

        public string GetBody(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].Body;
        }

        public string GetHtmlBody(string entryID)
        {
            if (!openMessages.ContainsKey(entryID))
            {
                return null;
            }

            return openMessages[entryID].HTMLBody;
        }
    }
}

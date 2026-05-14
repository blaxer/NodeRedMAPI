const edge = global.get('edge');

const getTodaysMessages = edge.func(function () {/*
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    public class Startup
    {
        public async Task<object> Invoke(dynamic input)
        {
            var results = new List<object>();

            Type outlookType = null;
            dynamic outlook = null;
            dynamic mapi = null;
            dynamic inbox = null;
            dynamic p5Folder = null;
            dynamic items = null;

            try
            {
                // Connect to Outlook
                outlookType = Type.GetTypeFromProgID("Outlook.Application");

                outlook = Activator.CreateInstance(outlookType);

                mapi = outlook.GetNamespace("MAPI");

                // Inbox
                inbox = mapi.GetDefaultFolder(6);          
                
                items = inbox.Items;

                // Sort newest first
                items.Sort("[ReceivedTime]", true);

                DateTime today = DateTime.Today;
                DateTime tomorrow = today.AddDays(1);

                int count = items.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic mail = items[i];

                    try
                    {
                        DateTime received =
                            (DateTime)mail.ReceivedTime;

                        // Only today's emails
                        if (received >= today &&
                            received < tomorrow)
                        {
                            string headers = "";

                            try
                            {
                                // PR_TRANSPORT_MESSAGE_HEADERS
                                headers =
                                    mail.PropertyAccessor.GetProperty(
                                        "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
                                    );
                            }
                            catch
                            {
                                headers = "";
                            }

                            results.Add(new
                            {
                                Subject = (string)mail.Subject,
                                Sender = (string)mail.SenderName,
                                SenderEmail = (string)mail.SenderEmailAddress,
                                To = (string)mail.To,
                                CC = (string)mail.CC,
                                Received = received.ToString("o"),

                                Headers = headers,

                                Body = (string)mail.Body,

                                EntryId = (string)mail.EntryID
                            });
                        }
                    }
                    catch
                    {
                        // Skip non-mail items
                    }
                }

                return results;
            }
            finally
            {
                if (items != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(items);

                if (p5Folder != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(p5Folder);

                if (inbox != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);

                if (mapi != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mapi);

                if (outlook != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlook);
            }
        }
    }
*/});

getTodaysMessages(null, function (err, result) {

    if (err) {
        node.error(err);
        return;
    }

    msg.payload = result;

    node.send(msg);
});

return;

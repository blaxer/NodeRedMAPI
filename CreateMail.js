const edge = global.get('edge');

const createInboxMessage = edge.func(function () {/*
    using System;
    using System.Threading.Tasks;

    public class Startup
    {
        public async Task<object> Invoke(dynamic input)
        {
            Type outlookType = null;

            dynamic outlook = null;
            dynamic mapi = null;

            dynamic mail = null;
            dynamic inbox = null;
            dynamic movedMail = null;

            try
            {
                // Outlook COM
                outlookType =
                    Type.GetTypeFromProgID("Outlook.Application");

                outlook =
                    Activator.CreateInstance(outlookType);

                mapi =
                    outlook.GetNamespace("MAPI");

                // Create mail item
                // 0 = olMailItem
                mail =
                    outlook.CreateItem(0);

                // Message fields
                mail.Subject =
                    input.Subject != null
                        ? (string)input.Subject
                        : "";

                mail.To =
                    input.To != null
                        ? (string)input.To
                        : "";

                // Body
                if (input.HtmlBody != null)
                {
                    mail.HTMLBody =
                        (string)input.HtmlBody;
                }
                else if (input.Body != null)
                {
                    mail.Body =
                        (string)input.Body;
                }

                // Fake sender display name
                if (input.SenderName != null)
                {
                    try
                    {
                        mail.SentOnBehalfOfName =
                            (string)input.SenderName;
                    }
                    catch
                    {
                    }
                }

               

                // Save first
                mail.Save();

                // Inbox = 6
                inbox =
                    mapi.GetDefaultFolder(6);

                // Move into Inbox
                movedMail =
                    mail.Move(inbox);

                // Mark unread
                movedMail.UnRead = true;

                // Save changes
                movedMail.Save();

                return new
                {
                    Success = true,
                    Subject = (string)movedMail.Subject,
                    EntryId = (string)movedMail.EntryID
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    Success = false,
                    Error = ex.ToString()
                };
            }
            finally
            {
                if (movedMail != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(movedMail);

                if (inbox != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);

                if (mail != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);

                if (mapi != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mapi);

                if (outlook != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlook);
            }
        }
    }
*/});

// USE msg.payload AS INPUT
createInboxMessage(msg.payload, function (err, result)
{
    if (err)
    {
        node.error(err);
        return;
    }

    msg.payload = result;

    node.send(msg);
});

return;



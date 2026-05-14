const edge = global.get('edge');

const moveMessage = edge.func(function () {/*
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
            dynamic targetFolder = null;
            dynamic movedMail = null;

            try
            {
                string entryId = (string)input.EntryId;
                string folderName = (string)input.TargetFolder;

                // Connect to Outlook
                outlookType = Type.GetTypeFromProgID("Outlook.Application");

                outlook = Activator.CreateInstance(outlookType);

                mapi = outlook.GetNamespace("MAPI");

                // Re-open message from EntryID
                mail = mapi.GetItemFromID(entryId);

                if (mail == null)
                {
                    throw new Exception("Mail item not found.");
                }

                // Inbox
                inbox = mapi.GetDefaultFolder(6);

                // Destination folder under Inbox
                targetFolder = inbox.Folders[folderName];
                

                if (targetFolder == null)
                {
                    throw new Exception(
                        "Target folder not found: " + folderName
                    );
                }

                // Move message
                movedMail = mail.Move(targetFolder);

                return new
                {
                    Success = true,
                    OriginalEntryId = entryId,
                    NewEntryId = movedMail.EntryID,
                    TargetFolder = folderName,
                    Subject = movedMail.Subject
                };
            }
            finally
            {
                if (movedMail != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(movedMail);

                if (targetFolder != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetFolder);

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

moveMessage(msg.payload, function (err, result) {

    if (err) {
        node.error(err);
        return;
    }

    msg.payload = result;

    node.send(msg);
});

return;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Threading.Tasks;

namespace Utility.Email
{
    class Emailer
    {
        private string clientServer;
        private string senderEmail;
        private string senderName;
        private string[] toRecipients;
        private string[] ccRecipients;
        private string localPath;

        public Emailer(string client_server, string sender_email, string sender_name, string to_recips, string cc_recips, string local_path = null)
        {
            clientServer = client_server;
            senderEmail = sender_email;
            senderName = sender_name;
            toRecipients = to_recips.Split(',');
            ccRecipients = cc_recips.Split(',');
            localPath = local_path;    
        }

        /// <summary>
        /// Constructs an email message and sends it via smtp to specified recipients
        /// </summary>
        /// /// <param name="subject">
        /// title of email message
        /// </param>
        /// <param name="body">
        /// content body of email
        /// </param>
        /// <param name="attachment">
        /// optional: filename/path of email attachment
        /// </param>
        public void Email(string subject, string body, string[] attachments=null)
        {
            // handle to smtp client object with specified SMTP host
            SmtpClient client = new SmtpClient(clientServer);
            // specify the message content.
            MailMessage message = new MailMessage();
            // add from recipients
            message.From = new MailAddress(senderEmail, senderName, System.Text.Encoding.UTF8);
            // add to recipients
            foreach (string toRecipient in toRecipients)
            {
                message.To.Add(toRecipient.Trim());
            }
            // add carbon copy recipients
            foreach (string ccRecipient in ccRecipients)
            {
                message.CC.Add(ccRecipient.Trim());
            }

            // add attachment to mail
            if (attachments != null)
            {
                foreach (string attachment in attachments)
                {
                    message.Attachments.Add(new Attachment(attachment));
                }
            }

            // include body of email
            message.Body = body;
            message.IsBodyHtml = true;
            // include subject of email
            message.Subject = subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            // construct local copy of email message
            // no handle to rename file. If desired, investigate FileSystemWatcher library
            if (localPath != null)
            {
                client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                client.PickupDirectoryLocation = Path.GetFullPath(localPath);
            }

            client.Send(message);

            // send email message over smtp network
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Send(message);

            // clean message object once send is complete
            message.Dispose();
        }
    }
}

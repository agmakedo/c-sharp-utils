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
        // handle to smtp client object with specified SMTP host
        private SmtpClient client;
        // specify the message content.
        private MailMessage message;

        private string clientServer;
        private string senderEmail;
        private string senderName;
        private string[] toRecipients;
        private string[] ccRecipients;
        private string localPath;

        /// <summary>
        /// Constructor for Emailer class
        /// </summary>
        /// /// <param name="_clientServer">
        /// SMTP server
        /// </param>
        /// <param name="_senderEmail">
        /// Email Address of sender
        /// </param>
        /// <param name="_senderName">
        /// Sender's email address alias
        /// </param>
        /// <param name="_toRecipients">
        /// Array of email addresses to be receiving message
        /// </param>
        /// <param name="_ccRecipients">
        /// Array of email addresses to be Carbon-Copied on email message
        /// </param>
        /// <param name="_localPath">
        /// Path to store email file locally before sending to SMTP server
        /// </param>
        public Emailer(string _clientServer, string _senderEmail, string _senderName, 
            string _toRecipients, string _ccRecipients, string _localPath = null)
        {
            clientServer = _clientServer;
            senderEmail = _senderEmail;
            senderName = _senderName;
            toRecipients = (_toRecipients.Length > 0) ? _toRecipients.Split(',') : null;
            ccRecipients = (_ccRecipients.Length > 0) ? _ccRecipients.Split(',') : null;
            localPath = _localPath;                     
        }

        private void SetEmailHeader()
        {
            // handle to smtp client object with specified SMTP host
            client = new SmtpClient(clientServer);
            // specify the message content.
            message = new MailMessage();

            // add from recipients
            message.From = new MailAddress(senderEmail, senderName, System.Text.Encoding.UTF8);
            // add to recipients
            if (toRecipients != null)
            {
                foreach (string toRecipient in toRecipients)
                {
                    message.To.Add(toRecipient.Trim());
                }
            }
            // add carbon copy recipients
            if (ccRecipients != null)
            {
                foreach (string ccRecipient in ccRecipients)
                {
                    message.CC.Add(ccRecipient.Trim());
                }
            }
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
        /// <param name="attachments">
        /// optional: filename/path of email attachment(s)
        /// </param>
        public void Email(string subject, string body, string[] attachments=null)
        {
            SetEmailHeader();

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

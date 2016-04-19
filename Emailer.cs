using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Threading.Tasks;

namespace RTUHM.Utility
{
    class Emailer
    {
        public Emailer()
        {
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
            SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["SMTPServer"]);
            // specify who is sending email message
            MailAddress from = new MailAddress("RTUHM@sempra.com",
                                               "RTU Health Monitor",
                                               System.Text.Encoding.UTF8);
            // specify who is receiving email message
            string[] toAddresses = ConfigurationManager.AppSettings["ToRecipients"].Split(',');
            // specify carbon copy recipients
            string[] ccAddresses = ConfigurationManager.AppSettings["CcRecipients"].Split(',');


            // specify the message content.
            MailMessage message = new MailMessage();
            // add from recipients
            message.From = new MailAddress("RTUHM@sempra.com","RTU Health Monitor", System.Text.Encoding.UTF8);
            // add to recipients
            foreach (string toAddress in toAddresses)
            {
                message.To.Add(toAddress.Trim());
            }
            // add carbon copy recipients
            foreach (string ccAddress in ccAddresses)
            {
                message.CC.Add(ccAddress.Trim());
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
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            client.PickupDirectoryLocation = Path.GetFullPath(ConfigurationManager.AppSettings["RTUHMReportFolderPath"]);

            client.Send(message);

            // send email message over smtp network
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Send(message);

            // clean message object once send is complete
            message.Dispose();
        }
    }
}

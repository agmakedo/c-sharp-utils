using System;
using System.Net.Mail;
using System.IO;

namespace Utility.Email
{
    public class Email
    {
        // handle to smtp client object with specified SMTP host
        private SmtpClient smtpClient;
        // specify the message content.
        private MailMessage mailMessage;

        private readonly string smtpServer;
        private readonly string sender;
        private readonly string senderAlias;
        private readonly string[] toRecipients;
        private readonly string[] ccRecipients;
        private readonly string localPath;

        private Email(string smtpServer, string sender, string senderAlias,
                      string[] toRecipients, string[] ccRecipients, string localPath)
        {
            this.smtpServer = smtpServer;
            this.sender = sender;
            this.senderAlias = senderAlias;
            this.toRecipients = toRecipients;
            this.ccRecipients = ccRecipients;
            this.localPath = localPath;

            ConfigureMailServer();
            ConfigureMailMessage();
        }

        private void ConfigureMailServer()
        {
            // handle to smtp client object with specified SMTP host
            smtpClient = new SmtpClient(smtpServer);          
        }

        private void ConfigureMailMessage()
        {
            try
            {
                // specify the message content.
                mailMessage = new MailMessage();

                // specify from recipient
                mailMessage.From = new MailAddress(sender, senderAlias, System.Text.Encoding.UTF8);
                
                // specify To recipients
                if (toRecipients != null)
                {
                    foreach (string toRecipient in toRecipients)
                    {
                        mailMessage.To.Add(toRecipient);
                    }
                }
                
                // specify carbon copy recipients
                if (ccRecipients != null)
                {
                    foreach (string ccRecipient in ccRecipients)
                    {
                        mailMessage.CC.Add(ccRecipient);
                    }
                }
            }
            catch(Exception ex)
            {                
                throw new Exception("Unable to construct email: " + ex.Message, ex.InnerException);
            }
        }

        /// <summary>
        /// Constructs an email message and sends it via smtp to specified recipients
        /// </summary>
        /// /// <param name="subject">
        /// title of email message
        /// </param>
        /// <param name="body">
        /// body of email
        /// </param>
        /// <param name="attachments">
        /// optional: filename/path of email attachment(s)
        /// </param>
        public void SendMessage(string subject, string body, string[] attachments = null)
        {
            // include subject of email
            mailMessage.Subject = subject;
            mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;

            // include body of email
            mailMessage.Body = body;
            mailMessage.IsBodyHtml = true;

            // include attachments to email
            if (attachments != null)
            {
                foreach (string attachment in attachments)
                {
                    mailMessage.Attachments.Add(new Attachment(attachment));
                }
            }
            
            // construct local copy of email message
            // no handle to rename file. If desired, investigate FileSystemWatcher library
            if (localPath != null)
            {
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                smtpClient.PickupDirectoryLocation = Path.GetFullPath(localPath);
                smtpClient.Send(mailMessage);
            }

            // send email message over smtp network
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Send(mailMessage);

            // clean message object once send is complete            
            mailMessage.Dispose();                        
        }

        /// <summary>
        /// Builder object allows user to construct customizable email class  
        /// </summary>
        public class ComposeHeader
        {
            private string smtpServer;
            private string sender;
            private string senderAlias;
            private string[] toRecipients;
            private string[] ccRecipients;
            private string localPath;

            public ComposeHeader() { }

            public ComposeHeader SMTPServer(string smtpServer)
            {
                this.smtpServer = smtpServer;
                return this;
            }

            public ComposeHeader Sender(string sender)
            {
                this.sender = sender;
                return this;
            }

            public ComposeHeader Sender(string sender, string senderAlias)
            {
                this.sender = sender;
                this.senderAlias = senderAlias;
                return this;
            }

            public ComposeHeader To(string[] toRecipients)
            {
                this.toRecipients = toRecipients;
                return this;
            }

            public ComposeHeader Cc(string[] ccRecipients)
            {
                this.ccRecipients = ccRecipients;
                return this;
            }

            public ComposeHeader StoreLocalCopy(string localPath)
            {
                this.localPath = localPath;
                return this;
            }

            public Email Bind()
            {
                return new Email(smtpServer, sender, senderAlias,
                                 toRecipients, ccRecipients, localPath);
            }
        }      
    }
}

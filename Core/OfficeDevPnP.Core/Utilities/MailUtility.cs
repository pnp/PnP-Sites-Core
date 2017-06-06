using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Security;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    public class MailUtility
    {
        /// <summary>
        /// Sends an email via Office 365 SMTP
        /// </summary>
        /// <param name="servername">Office 365 SMTP address. By default this is smtp.office365.com.</param>
        /// <param name="fromAddress">The user setting up the SMTP connection. This user must have an EXO license.</param>
        /// <param name="fromUserPassword">The password of the user.</param>
        /// <param name="to">List of TO addresses.</param>
        /// <param name="cc">List of CC addresses.</param>
        /// <param name="subject">Subject of the mail.</param>
        /// <param name="body">HTML body of the mail.</param>
        /// <param name="sendAsync">Sends the email asynchronous so as to not block the current thread (default: false).</param>
        /// <param name="asyncUserToken">The user token that is used to correlate the asynchronous email message.</param>
        public static void SendEmail(string servername, string fromAddress, string fromUserPassword, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body, bool sendAsync = false, object asyncUserToken = null)
        {
            // Get the secure password
            var secureString = new SecureString();
            foreach (char c in fromUserPassword.ToCharArray())
            {
                secureString.AppendChar(c);
            }

            SendEmail(servername, fromAddress, secureString, to, cc, subject, body, sendAsync, asyncUserToken);
        }

        /// <summary>
        /// Sends an email via Office 365 SMTP
        /// </summary>
        /// <param name="servername">Office 365 SMTP address. By default this is smtp.office365.com.</param>
        /// <param name="fromAddress">The user setting up the SMTP connection. This user must have an EXO license.</param>
        /// <param name="fromUserPassword">The secure password of the user.</param>
        /// <param name="to">List of TO addresses.</param>
        /// <param name="cc">List of CC addresses.</param>
        /// <param name="subject">Subject of the mail.</param>
        /// <param name="body">HTML body of the mail.</param>
        /// <param name="sendAsync">Sends the email asynchronous so as to not block the current thread (default: false).</param>
        /// <param name="asyncUserToken">The user token that is used to correlate the asynchronous email message.</param>
        public static void SendEmail(string servername, string fromAddress, SecureString fromUserPassword, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body, bool sendAsync = false, object asyncUserToken = null)
        {
            SmtpClient client = CreateSmtpClient(servername, fromAddress, fromUserPassword);
            MailMessage mail = CreateMailMessage(fromAddress, to, cc, subject, body);
            try
            {
                if (sendAsync)
                {
                    client.SendCompleted += (sender, args) =>
                    {
                        if (args.Error != null)
                        {
                            Diagnostics.Log.Error(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendFailed, args.Error.Message);
                        }
                        else if (args.Cancelled)
                        {
                            Diagnostics.Log.Info(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendMailCancelled);
                        }
                    };
                    client.SendAsync(mail, asyncUserToken);
                }
                else
                {
                    client.Send(mail);
                }
            }
            catch (SmtpException smtpEx)
            {
                Diagnostics.Log.Error(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendException, smtpEx.Message);
            }
            catch (Exception ex)
            {
                Diagnostics.Log.Error(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendExceptionRethrow0, ex);
                throw;
            }
        }

        /// <summary>
        /// Sends an email via Office 365 SMTP as an asynchronous operation
        /// </summary>
        /// <param name="servername">Office 365 SMTP address. By default this is smtp.office365.com.</param>
        /// <param name="fromAddress">The user setting up the SMTP connection. This user must have an EXO license.</param>
        /// <param name="fromUserPassword">The password of the user.</param>
        /// <param name="to">List of TO addresses.</param>
        /// <param name="cc">List of CC addresses.</param>
        /// <param name="subject">Subject of the mail.</param>
        /// <param name="body">HTML body of the mail.</param>
        public static async Task SendEmailAsync(string servername, string fromAddress, string fromUserPassword, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body)
        {
            // Get the secure password
            var secureString = new SecureString();
            foreach (char c in fromUserPassword.ToCharArray())
            {
                secureString.AppendChar(c);
            }

            await SendEmailAsync(servername, fromAddress, secureString, to, cc, subject, body);
        }

        /// <summary>
        /// Sends an email via Office 365 SMTP as an asynchronous operation
        /// </summary>
        /// <param name="servername">Office 365 SMTP address. By default this is smtp.office365.com.</param>
        /// <param name="fromAddress">The user setting up the SMTP connection. This user must have an EXO license.</param>
        /// <param name="fromUserPassword">The secure password of the user.</param>
        /// <param name="to">List of TO addresses.</param>
        /// <param name="cc">List of CC addresses.</param>
        /// <param name="subject">Subject of the mail.</param>
        /// <param name="body">HTML body of the mail.</param>
        public static async Task SendEmailAsync(string servername, string fromAddress, SecureString fromUserPassword, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body)
        {
            SmtpClient client = CreateSmtpClient(servername, fromAddress, fromUserPassword);
            MailMessage mail = CreateMailMessage(fromAddress, to, cc, subject, body);
            try
            {
                await client.SendMailAsync(mail);
            }
            catch (SmtpException smtpEx)
            {
                Diagnostics.Log.Error(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendException, smtpEx.Message);
            }
            catch (Exception ex)
            {
                Diagnostics.Log.Error(Constants.LOGGING_SOURCE, CoreResources.MailUtility_SendExceptionRethrow0, ex);
                throw;
            }
        }

        /// <summary>
        /// Sends an email using the SharePoint SendEmail method
        /// </summary>
        /// <param name="context">Context for SharePoint objects and operations</param>
        /// <param name="to">List of TO addresses.</param>
        /// <param name="cc">List of CC addresses.</param>
        /// <param name="subject">Subject of the mail.</param>
        /// <param name="body">HTML body of the mail.</param>
        public static void SendEmail(ClientContext context, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body)
        {
            EmailProperties properties = new EmailProperties();
            properties.To = to;

            if (cc != null)
            {
                properties.CC = cc;
            }

            properties.Subject = subject;
            properties.Body = body;

            Microsoft.SharePoint.Client.Utilities.Utility.SendEmail(context, properties);
            context.ExecuteQueryRetry();
        }

        private static SmtpClient CreateSmtpClient(string serverName, string fromAddress, SecureString fromUserPassword)
        {
            if (String.IsNullOrEmpty(serverName))
            {
                throw new ArgumentException("serverName");
            }

            if (String.IsNullOrEmpty(fromAddress))
            {
                throw new ArgumentException("fromAddress");
            }

            if (fromUserPassword == null || fromUserPassword.Length == 0)
            {
                throw new ArgumentException("fromUserPassword");
            }

            return new SmtpClient(serverName)
            {
                Port = 587,
                EnableSsl = true,
                Credentials = new NetworkCredential(fromAddress, fromUserPassword)
            };
        }

        private static MailMessage CreateMailMessage(string fromAddress, IEnumerable<String> to, IEnumerable<String> cc, string subject, string body)
        {
            MailMessage mail = new MailMessage()
            {
                From = new MailAddress(fromAddress),
                Subject = subject,
                Body = body,
                IsBodyHtml = true
            };

            foreach (string user in to)
            {
                mail.To.Add(user);
            }

            if (cc != null)
            {
                foreach (string user in cc)
                {
                    mail.CC.Add(user);
                }
            }

            return mail;
        }
    }
}
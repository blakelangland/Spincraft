using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Net.Mail;

using Ice.Core;
using Erp.BO;
using Erp.Proxy.BO;
using Ice.Lib.Framework;
using Ice.Adapters;
using Ice.BO;


namespace HorizonScientific
{
    public class HSEmailHelper
    {
        #region Constructors
        public HSEmailHelper()
        {
        }
        #endregion

        #region Methods

        public static void SendEmail(List<string> oToAddresses, string sSubject, string sBody, List<string> oAttachments = null)
        {
            try
            {
                MailMessage oMailMessage = new MailMessage();
                MailAddress oFrom = new MailAddress(UserId);
                oMailMessage.From = oFrom;

                foreach (string sTo in oToAddresses)
                {
                    MailAddress oTo = new MailAddress(sTo);
                    oMailMessage.To.Add(oTo);
                }

                oMailMessage.Subject = sSubject;

                oMailMessage.Body = sBody;

                if (oAttachments != null)
                {
                    foreach (string sFileName in oAttachments)
                    {
                        oMailMessage.Attachments.Add(new Attachment(sFileName));
                    }
                }
                SmtpClient oClient = new SmtpClient(Host, Port);
                oClient.Credentials = new System.Net.NetworkCredential(UserId, Password);
                oClient.EnableSsl = false;
                oClient.Send(oMailMessage);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        #endregion

        #region Properties

        public static string Host
        {
            get
            {
                return g_sHost;
            }
        }

        public static int Port
        {
            get
            {
                return g_iPort;
            }
        }

        public static string UserId
        {
            get
            {
                return HSCredentialStore.Get(HSCredentialStore.Keys.NoReplyLogin);
            }
        }

        public static string Password
        {
            get
            {
                return HSCredentialStore.Get(HSCredentialStore.Keys.NoReplyPassword);
            }
        }

        #endregion

        #region Data Members

        private static string g_sHost = "10.120.50.25";
        private static int g_iPort = 25;

        #endregion
    }
}

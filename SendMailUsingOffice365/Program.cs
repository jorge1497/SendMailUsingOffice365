using System;
using System.Net.Mail;

namespace SendMailUsingOffice365
{
    class Program
    {
        static void Main(string[] args)
        {
            //Fill parameters in MailMessage object
            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress("someone@someone.com", "Someone_name"));
            msg.From = new MailAddress("Your_office365_mail@Your_domain.com", "Your_name");
            msg.Subject = "This is a Test Mail";
            msg.Body = "This is a test message body";
            msg.IsBodyHtml = true;

            //Set Credentials in SmtpClient object
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("Your_office365_mail@Your_domain.com", "Your_password");
            client.Port = 25;
            client.Host = "smtp.office365.com";
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = true;
            try
            {
                //Send Mail
                client.Send(msg);
                Console.Write("Message Sent Succesfully");
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }
        }
    }
}

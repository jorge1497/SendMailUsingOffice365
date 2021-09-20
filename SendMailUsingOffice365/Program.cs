using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

namespace SendMailUsingOffice365
{
    class Program
    {
        static void Main(string[] args)
        {
            //attachments filepath
            var file1 = Path.Combine(Directory.GetCurrentDirectory(), "E:\\CareOregon\\attachments\\test1.txt");
            var file2 = Path.Combine(Directory.GetCurrentDirectory(), "E:\\CareOregon\\attachments\\test2.txt");


            //Fill parameters in MailMessage object
            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress("swap407@gmail.com", "Swapnil"));
            msg.From = new MailAddress("swapnil.sonawane@excellarate.com", "Swapnil");
            msg.Subject = "This is a Test Mail";
            //msg.Body = "This is a test message body";

            //add multiple attachments here
            msg.Attachments.Add(new Attachment(file1));
            msg.Attachments.Add(new Attachment(file2));
            msg.IsBodyHtml = true;

            //add inline images and body text
            msg.AlternateViews.Add(GetMailBody());


            //Set Credentials in SmtpClient object
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("swapnil.sonawane@excellarate.com", "@swap1212");
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

        static private AlternateView GetMailBody()
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "E:\\CareOregon\\attachments\\rose.jpg");
            LinkedResource Img = new LinkedResource(path, MediaTypeNames.Image.Jpeg);
            Img.ContentId = "MyImage";
            string str = @"  
            <table>  
                <tr>  
                    <td> This is a test message body </td>  
                </tr>  
                <tr>  
                    <td>  
                      <img src=cid:MyImage  id='img' alt='' width='100px' height='100px'/>   
                    </td>  
                </tr>
                <tr>  
                    <td> This is a test message body </td>  
                </tr>
            </table>  
            ";
            AlternateView AV = AlternateView.CreateAlternateViewFromString(str, null, MediaTypeNames.Text.Html);
            AV.LinkedResources.Add(Img);
            return AV;
        }
    }
}

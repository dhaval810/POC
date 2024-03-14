using MimeKit;

namespace EmailFunction
{
    public class EmailService
    {
        public static void SendEmail()
        {
            var client = new SmtpClient("sandbox.smtp.mailtrap.io", 2525)
            {
                Credentials = new NetworkCredential("ef498c8fb0e515", "221fd41f33c1bb"),
                EnableSsl = true
            };

            try
            {
                client.Send("from@example.com", "dhavalprajapati81097@gmail.com", "Hello world", "testbody");
                // Sending email
                //smtpClient.Send(mailMessage);
                Console.WriteLine("Email sent successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}

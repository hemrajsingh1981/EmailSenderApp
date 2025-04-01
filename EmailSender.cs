using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using CsvHelper;
using Newtonsoft.Json;

public class Config
{
    public string SmtpServer { get; set; }
    public int SmtpPort { get; set; }
    public string SmtpUsername { get; set; }
    public string SmtpPassword { get; set; }
    public string SenderEmail { get; set; }
    public string Subject { get; set; }
    public string HtmlTemplate { get; set; }
    public List<string> EmbeddedImages { get; set; }
    public int BatchSize { get; set; }
}

public class Recipient
{
    public string Email { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
}

public class EmailSender
{
    public static void Main()
    {
        try
        {
            // ✅ Load configuration
            string configContent = File.ReadAllText("config.json");
            Config config = JsonConvert.DeserializeObject<Config>(configContent);

            // ✅ Read recipients from CSV
            List<Recipient> recipients;
            using (var reader = new StreamReader("recipients.csv"))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                recipients = csv.GetRecords<Recipient>().ToList();
            }

            // ✅ Read the HTML email template
            string emailBody = File.ReadAllText(config.HtmlTemplate);

            // ✅ Setup SMTP Client
            SmtpClient client = new SmtpClient(config.SmtpServer, config.SmtpPort)
            {
                Credentials = new NetworkCredential(config.SmtpUsername, config.SmtpPassword),
                EnableSsl = true
            };

            // ✅ Send emails in batches
            int batchSize = config.BatchSize;
            for (int i = 0; i < recipients.Count; i += batchSize)
            {
                List<Recipient> batch = recipients.Skip(i).Take(batchSize).ToList();
                Console.WriteLine($"📩 Sending batch {i / batchSize + 1} of {batchSize} emails...");

                foreach (var recipient in batch)
                {
                    int retryCount = 0;
                    bool sentSuccessfully = false;

                    while (retryCount < 3 && !sentSuccessfully)
                    {
                        try
                        {
                            // Personalize email content
                            string personalizedBody = emailBody
                                .Replace("{first_name}", recipient.FirstName)
                                .Replace("{last_name}", recipient.LastName);

                            // Create email message
                            MailMessage mail = new MailMessage
                            {
                                From = new MailAddress(config.SenderEmail),
                                Subject = config.Subject,
                                IsBodyHtml = true
                            };
                            mail.To.Add(new MailAddress(recipient.Email.Trim()));

                            // ✅ Embed images
                            AlternateView av = AlternateView.CreateAlternateViewFromString(personalizedBody, null, "text/html");
                            foreach (var imagePath in config.EmbeddedImages)
                            {
                                LinkedResource image = new LinkedResource(imagePath, "image/png")
                                {
                                    ContentId = Path.GetFileNameWithoutExtension(imagePath)
                                };
                                av.LinkedResources.Add(image);
                            }
                            mail.AlternateViews.Add(av);

                            // ✅ Attach files from 'attachments' folder
                            string attachmentsFolder = "attachments";
                            if (Directory.Exists(attachmentsFolder))
                            {
                                foreach (var file in Directory.GetFiles(attachmentsFolder))
                                {
                                    mail.Attachments.Add(new Attachment(file));
                                }
                            }

                            // ✅ Send email
                            client.Send(mail);
                            sentSuccessfully = true;
                            LogSuccessEmail(recipient.Email);
                            Console.WriteLine($"✅ Email sent successfully to {recipient.Email}");
                        }
                        catch (Exception ex)
                        {
                            retryCount++;
                            Console.WriteLine($"⚠️ Failed to send email to {recipient.Email} (Attempt {retryCount}/3): {ex.Message}");

                            if (retryCount == 3)
                            {
                                LogFailedEmail(recipient.Email, ex.Message);
                            }
                            else
                            {
                                Thread.Sleep(2000); // Wait before retrying
                            }
                        }
                    }
                }

                // ✅ Pause to prevent rate limits
                Thread.Sleep(3000);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }

    private static void LogSuccessEmail(string email)
    {
        string logMessage = $"{DateTime.Now}: Successfully sent email to {email}\n";
        File.AppendAllText("log.txt", logMessage);
        Console.WriteLine(logMessage);
    }

    private static void LogFailedEmail(string email, string error)
    {
        string logMessage = $"{DateTime.Now}: Failed to send email to {email}. Error: {error}\n";
        File.AppendAllText("failed_emails.txt", logMessage);
        Console.WriteLine(logMessage);
    }
}

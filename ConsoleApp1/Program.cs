using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using TemplateEngine.Docx;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string server = "smtp.yandex.ru";
            string add = "address@yandex.ru";

            SmtpClient client = new SmtpClient(server); // используемый сервер
            client.Credentials = new NetworkCredential(add, "somePassword"); // имя пользователя и пароль для почты
            client.EnableSsl = true;

            MailMessage message = new MailMessage(add, add); // письмо (от кого, кому)


            byte[] byteArray = File.ReadAllBytes("Docs\\InputTemplate.docx");  // документ исходник

            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);

                int studsN = 164;
                int studsGr = 14;

                //File.Delete("Docs\\OutputDocument.html");
                //File.Copy("Docs\\InputTemplate.docx", "Docs\\OutputDocument.html");

                var valuesToFill = new Content(
                    new FieldContent("Report date", DateTime.Now.ToString())); // заполнение тегов

                var newStudents = new Content(
                    new FieldContent("New Students", studsN.ToString()));

                var gradStudents = new Content(
                    new FieldContent("Graduating students", studsGr.ToString()));

                var college = new Content(
                    new FieldContent("College", "RCM systems"));

                var difference = new Content(
                    new FieldContent("Difference", ((studsN - studsGr).ToString())));


                using (var outputDocument = new TemplateProcessor(memoryStream) 
                    .SetRemoveContentControls(true))
                {
                    outputDocument.FillContent(valuesToFill); // заполняем значения тегов в документе
                    outputDocument.FillContent(newStudents);
                    outputDocument.FillContent(gradStudents);
                    outputDocument.FillContent(college);
                    outputDocument.FillContent(difference);
                    outputDocument.SaveChanges();
                }

                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true)) // сейчас превратим в html документ
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title"
                    };

                    message.Subject = "Using the new SMTP client.";

                    XElement html = HtmlConverter.ConvertToHtml(doc, settings); // прекращение

                    var reader = html.CreateReader();
                    reader.MoveToContent();

                    message.Body = reader.ReadInnerXml(); // запихиваем в письмо
                    message.IsBodyHtml = true; // обязательно указываем что у нас там html

                    //File.WriteAllText("Docs\\OutputDocument.html", html.ToStringNewLineOnAttributes()); // если надо, то сохраняем в файл
                }
            }

            try
            {
                client.Send(message); // отправляем письмо
            }
            catch (Exception ex) // ловим ошибки, если накосячили при отправке письма
            {
                Console.WriteLine("Exception caught in CreateTestMessage2(): {0}",
                    ex.ToString());
            }
        }
    }
}

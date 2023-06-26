using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace ServiceProject
{
    public partial class Service1 : ServiceBase
    {
        public Timer ScheduleTimer;

        List<string> recipientEmails = new List<string>
        {
            "panchalmeet1302@gmail.com",
            "panchalpriti714@gmail.com",
            "npsmtp217@gmail.com"
        };

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Thread threadStart = new Thread(new ThreadStart(ScheduleService));
            //this.ScheduleService();
            threadStart.Start();
        }

        public void ScheduleService()
        {
            try
            {
                DateTime now = DateTime.Now;
                DateTime scheduleTime = DateTime.Parse("09:40");

                if (now > scheduleTime)
                {
                    scheduleTime = scheduleTime.AddDays(1);
                }

                TimeSpan timeSpan = scheduleTime.Subtract(now);

                ScheduleTimer = new Timer(new TimerCallback(SendMail), null, timeSpan, TimeSpan.FromDays(1));
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex);
            }
        }

        public void SendMail(object e)
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("panchalmeet1302@gmail.com");
                    foreach (string email in recipientEmails)
                    {
                        mail.To.Add(email);
                    }

                    mail.Subject = "Reminder!";
                    mail.IsBodyHtml = true;
                    string htmlBody;
                    htmlBody = @"<html>
                      <body>
                      <h1>Reminder!!</h1>
                      <h3> Don't Forget to fill your attandance!! </h3>
                      </body>
                      </html>
                     ";
                    mail.Body = htmlBody;

                    using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                    {
                        smtp.Credentials = new NetworkCredential("meetpanchal194@gmail.com", "ksdqxndnbbsofpyz");
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                    this.ScheduleService();
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex);
            }
        }

        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter streamWriter = null;
            try
            {
                streamWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                streamWriter.WriteLine(DateTime.Now.ToString() + ":" + ex.Source.ToString().Trim() + ";" + ex.Message.ToString().Trim());
                streamWriter.Flush();
                streamWriter.Close();
            }
            catch
            {

            }
        }

        public static void WriteErrorLog(string message)
        {
            StreamWriter streamWriter = null;
            try
            {
                streamWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                streamWriter.WriteLine(DateTime.Now.ToString() + ":" + message);
                streamWriter.Flush();
                streamWriter.Close();
            }
            catch
            {

            }
        }

        protected override void OnStop()
        {
            ScheduleTimer.Dispose();
        }
    }
}

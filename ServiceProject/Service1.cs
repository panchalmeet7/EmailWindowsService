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

        #region User Emails Object
        List<string> recipientEmails = new List<string>
        {
            "panchalmeet1302@gmail.com",
            "panchalpriti714@gmail.com",
            "npsmtp217@gmail.com"
        };
        #endregion

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

        protected override void OnStop()
        {
            ScheduleTimer.Dispose();
        }

        /// <summary>
        /// This function compares scheduleTime with current datetime if it matches then and then SendMail function will get called
        /// </summary>
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

        #region Send Email Code Function
        /// <summary>
        ///  Send Email with given subject and message. 
        /// </summary>
        /// <param name="e"></param>
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
        #endregion


        /// <summary>  
        /// This function write log to LogFile.text when some error occurs.  
        /// </summary>  
        /// <param name="ex"></param>  
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

        /// <summary>  
        /// This function write Message to log file.  
        /// </summary>  
        /// <param name="Message"></param> 
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

       
    }
}

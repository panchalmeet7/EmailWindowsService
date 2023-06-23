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
                DateTime scheduleTime = DateTime.Parse("17:50");

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
                    mail.To.Add("meetpanchal194@gmail.com");
                    mail.Subject = "Reminder!";
                    mail.Body = "Oh hey there! Dont Forget to Fill the worklog in the workspace!";
                    
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

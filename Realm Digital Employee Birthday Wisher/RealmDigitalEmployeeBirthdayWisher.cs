using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Net.Http;

namespace Realm_Digital_Employee_Birthday_Wisher
{
    public partial class RealmDigitalEmployeeBirthdayWisher : ServiceBase
    {
        System.Timers.Timer createOrderTimer;
        public RealmDigitalEmployeeBirthdayWisher()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ToString());
                string selectCmd = "select id,name, lastname, dateOfBirth, employmentStartDate, employmentEndDate from Employees";
                SqlDataAdapter ad = new SqlDataAdapter(selectCmd, con);
                DataTable dt = new DataTable();

                try
                {
                    ad.Fill(dt);               
                }

                catch (Exception ex)
                {
                    WriteToLogFile(ex.Message);               
                }
                foreach (DataRow dr in dt.Rows)
                {
                    DateTime today = DateTime.Now;
                    DateTime dtDOB = (DateTime)dr["dob"];

                    if (Convert.ToDateTime(dr["dob"]) == today)
                    {
                        //if(DateTime.IsLeapYear)
                        //{  }
                        //else
                        //{ }
                        if (Convert.ToDateTime(dr["employmentStartDate"]) <= today && Convert.ToDateTime(dr["employmentEndDate"]) == today)
                        {
                            string dow = today.DayOfWeek.ToString().ToLower();
                            if (dow == "monday")
                            {
                                DateTime lastSunday = today.AddDays(-1);
                                DateTime lastSaturday = today.AddDays(-2);
                                if ((dtDOB.Day == today.Day || dtDOB.Day == lastSaturday.Day || dtDOB.Day == lastSunday.Day) && dtDOB.Month == today.Month)
                                {
                                    sendmail(dr);
                                }
                            }
                            else
                            {
                                if (dtDOB.Day == today.Day && dtDOB.Month == today.Month)
                                {
                                    sendmail(dr);
                                }
                            }
                        }
                    }
                }
                createOrderTimer = new System.Timers.Timer();                    
                createOrderTimer.Interval = Convert.ToDouble(ConfigurationManager.AppSettings["ServiceRunInterval"].ToString());
                createOrderTimer.Enabled = true;
                createOrderTimer.Elapsed += new System.Timers.ElapsedEventHandler(createOrderTimer_Elapsed);
            }
            catch(Exception ex)
            {
                WriteToLogFile(ex.Message);
            }
        }
        protected void createOrderTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs args)
        {
            createOrderTimer.Stop();

            ServiceController controller = new ServiceController("RealmDigitalEmployeeBirthdayWisher");
            
            controller.Stop();
        }
        public void sendmail(DataRow dr)
        {
            String mailServer = ConfigurationManager.AppSettings["MailServer"].ToString();
            String mailtxt = "";
            MailMessage mail = new MailMessage();
            mail.IsBodyHtml = true;
            mail.To.Add(dr["email"].ToString());
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFromAddress"].ToString());
            mail.Subject = "Happy Birthday";
            mailtxt = "<font face='verdana' color='#FF9900'><b>" + "Hi " + dr["name"].ToString()+" " + dr["lastname"].ToString() + "," + "</b></font><br><br>";
            mailtxt = mailtxt + "<font face='verdana' color='#FF0000'><b>" + "Wishing you a very HAPPY BIRTHDAY........and many more." + "</b></font><br><br>";
            mailtxt = mailtxt + "<font face='verdana' color='#008080'><b>" + "May today be filled with sunshine and smile, laughter and love." + "</b></font><br><br>";
            mailtxt = mailtxt + "<font face='verdana' color='#0000FF'><b>Cheers!" + "<br><br>";
            mail.Body = mailtxt;
            mail.Priority = MailPriority.Normal;
            SmtpClient client = new SmtpClient(mailServer);

            client.Send(mail);
        }
        //protected override void OnStop()
        //{
        //}
        public void WriteToLogFile(string content)
        {
            if (!(Directory.Exists(ConfigurationManager.AppSettings["LogFilePath"].ToString())))
            {
                Directory.CreateDirectory(ConfigurationManager.AppSettings["LogFilePath"].ToString());
            }

            StringBuilder strLogFilePath = new StringBuilder();
            strLogFilePath.Append(ConfigurationManager.AppSettings["LogFilePath"].ToString());
            strLogFilePath.Append(@"\LogFile.txt");

            if (!File.Exists(strLogFilePath.ToString()))
            {
                FileStream fsErrorLog = new FileStream(strLogFilePath.ToString(), FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fsErrorLog.Close();
            }

            StreamWriter writer = new StreamWriter(strLogFilePath.ToString(), true);
            writer.WriteLine("********************************************************************");
            writer.WriteLine("Date " + DateTime.Now);
            writer.WriteLine("Content: " + content);
            writer.WriteLine("********************************************************************");
            writer.Close();
        }
    }
}

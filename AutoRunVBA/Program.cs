using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Collections;
using System.Text;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace Macro
{

    public class Program
    {
        public static void Main(string[] args)
        {
            Program test = new Program();
            test.ExecuteMacro();
        }
        public class nameList
        {
            public string Name { get; set; }
            public string LastName { get; set; }
            public int Remaining { get; set; }
        }
        public void ExecuteMacro()
        {

            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsetting.json")
                .AddEnvironmentVariables()
                .Build();
            string strConnection = config.GetConnectionString("DbConnection");
            string MailSender = config.GetSection("Parameter").GetSection("Mailsender").Value;
            string SenderPass = config.GetSection("Parameter").GetSection("SenderPass").Value;
            string Mailreceiver = config.GetSection("Parameter").GetSection("Mailreceiver").Value;
            string workSheet = config.GetSection("Parameter").GetSection("workSheet").Value;

            string connString = "Provider= Microsoft.ACE.OLEDB.16.0;" + "Data Source=" + strConnection + ";Extended Properties='Excel 8.0;HDR=No'";
            // Create the connection object
            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + workSheet + "$]", oledbConn);
                // Create new OleDbDataAdapter
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();
                // Fill the DataSet from the data extracted from the worksheet.
                oleda.Fill(ds, "Employees");

                ArrayList arr = new ArrayList();
                List<nameList> nl = new List<nameList>();
                string[,] arrw;
                //loop through each row
                foreach (var m in ds.Tables[0].DefaultView)
                {
                    try
                    {

                        int num = (Convert.ToInt32(((System.Data.DataRowView)m).Row.ItemArray[9]));

                        if (num < 15)
                        {
                            string name = ((System.Data.DataRowView)m).Row.ItemArray[1].ToString();
                            string lastname = ((System.Data.DataRowView)m).Row.ItemArray[2].ToString();
                            nameList n = new nameList();
                            n.Name = name;
                            n.LastName = lastname;
                            n.Remaining = num;
                            nl.Add(n);
                        }

                    }
                    catch (Exception e)
                    {
                    }

                    oledbConn.Close();
                }
                //End Each loop

                        // Check exist data
                        if (nl.Count != 0)
                        {
                            Console.WriteLine("Not Null");
                
                            StringBuilder sb = new StringBuilder();

                            foreach (var obj in nl)
                            {
                                sb.AppendFormat("<br/>{0} {1}-san - Remaining Day: {2}", obj.Name, obj.LastName, obj.Remaining);
                            }

                            try
                            {
                                var networkCredential = new NetworkCredential(MailSender, SenderPass);
                                MailMessage mail = new MailMessage();
                                SmtpClient SmtpServer = new SmtpClient("smtp.office365.com");
                                SmtpServer.Credentials = networkCredential;
                                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                                mail.From = new MailAddress(@"" + MailSender + "");
                                mail.To.Add(@"" + Mailreceiver + "");
                                string subject = "Notify: Workpermit Remaining Day";
                                mail.Subject = subject;
                                string htmlBody = "Dear All<br/> <br/> Please check Japanease VISA and workpermit status follow as below: <br/>" + sb.ToString();
                                mail.Body = htmlBody;
                                mail.IsBodyHtml = true;
                                SmtpServer.Port = 587;
                                SmtpServer.EnableSsl = true;
                                SmtpServer.UseDefaultCredentials = false;
                                SmtpServer.Credentials = networkCredential;
                                SmtpServer.Send(mail);

                            }
                            catch (Exception ex) { }
                        }
                    }
                catch (Exception e)
            {
                Console.WriteLine("Error :" + e.Message);
            }
        
            finally
            {
                // Close connection
                oledbConn.Close();
            }
        }
    }
}
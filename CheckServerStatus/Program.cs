using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CheckServerStatus
{
    class Program
    {
        static void Main(string[] args)
        {
            string x = "ddd IP : 10.10.10.88 - AD01-DC fff";
            string[] y = x.Split(new string[] { "IP : 10.10.10.88 - AD01-DC" }, StringSplitOptions.None);
            

            var column = new List<data>();
            List<parameter> driveInfo = new List<parameter>();
            using (var rd = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "\\master.csv"))
            {
                while (!rd.EndOfStream)
                {
                    var splits = rd.ReadLine().Split(';');
                    if (splits[0] != "ip")
                    {
                        column.Add(new data
                        {
                            ip = splits[0],
                            user = splits[1],
                            pwd = splits[2]
                        });
                    }
                }
            }

            foreach (var item in column)
            {
                ConnectionOptions options;
                options = new ConnectionOptions();
                var server = item.ip;
                options.Username = item.user;
                options.Password = item.pwd;
                options.EnablePrivileges = true;
                options.Impersonation = ImpersonationLevel.Impersonate;
                Console.WriteLine(server);
                ManagementScope scope;
                scope = new ManagementScope("\\\\" + server + "\\root\\cimv2", options);
                scope.Connect();

                DateTime Now = DateTime.Now;


                SelectQuery query = new SelectQuery("select * from Win32_LogicalDisk where DriveType=3");

                //execute the query using WMI
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                //loop through each drive found
                foreach (ManagementObject drive in searcher.Get())
                {
                    driveInfo.Add(new parameter
                    {
                        ip = item.ip,
                        name = drive["name"].ToString(),
                        capacity = drive["FreeSpace"].ToString()
                    });
                    var xx = Convert.ToDouble(drive["FreeSpace"]) / 1024 / 1024 / 1024;
                    var yy = Convert.ToDouble(drive["Size"]) - Convert.ToDouble(drive["FreeSpace"]);
                }

            }

            List<string> msg = new List<string>();
            foreach (var item in column)
            {
                List<string> capacity = new List<string>();
                foreach (var item2 in driveInfo)
                {
                    if (item2.ip == item.ip)
                    {
                        capacity.Add(item2.name+"\nFree : "+item2.capacity);
                    }
                }
                msg.Add("IP : " + item.ip + "\n data : " + string.Join(",",capacity));
            }

        }
        public class data
        {
            public string ip { get; set; }
            public string user { get; set; }
            public string pwd { get; set; }
        }

        public class parameter
        {
            public string ip { get; set; }
            public string name { get; set; }
            public string capacity { get; set; }
        }
        //private static void sendFileTelegram(string chatId, string body)
        //{
        //    ServicePointManager.Expect100Continue = true;
        //    ServicePointManager.DefaultConnectionLimit = 9999;

        //    var client = new RestClient("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendDocument");
        //    RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendDocument", Method.POST);


        //    requestWa.Timeout = -1;
        //    requestWa.AddHeader("Content-Type", "multipart/form-data");
        //    requestWa.AddParameter("chat_id", chatId);
        //    requestWa.AddFile("document", body);
        //    var responseWa = client.ExecutePostAsync(requestWa);
        //    Console.WriteLine(responseWa.Result.Content);
        //}
    }
}

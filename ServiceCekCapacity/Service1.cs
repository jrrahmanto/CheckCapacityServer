using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ServiceCekCapacity
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            working();
            Timer timer = new Timer();
            timer.Interval = 60000;  //15 detik
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            var jam = DateTime.Now.Hour;
            var menit = DateTime.Now.Minute;
            if (jam == 8 && menit == 00 || jam == 16 && menit == 00 || jam == 21 && menit == 00)
            {
                working();
                monitoringServices("DTI Check Capacity Server", "I", "Check capacity server ok");
            }
            //string chat_id = "6281310215750-1589649694@g.us";

            //string file = AppDomain.CurrentDomain.BaseDirectory + "\\logLastMessage.txt";
            //string[] text = File.ReadAllLines(file);

            //string messages = GetMessageList(chat_id, Convert.ToInt32(text[0]));
            //ResponseChat rc = JsonConvert.DeserializeObject<ResponseChat>(messages);

            //foreach (MessageChat mc in rc.messages)
            //{
            //    string text2 = System.IO.File.ReadAllText(file);
            //    text2 = mc.messageNumber.ToString();
            //    System.IO.File.WriteAllText(file, text2);
            //    WriteToFile("1" + mc.body);
            //    if (mc.body.ToLower() == "check capacity server")
            //    {
            //        working();
            //    }
            //}
        }
        public void working()
        {
            string file = AppDomain.CurrentDomain.BaseDirectory + "\\numberregistered.txt";
            string[] text = File.ReadAllLines(file);
            List<string> number = new List<string>();
            foreach (string item in text)
            {
                string[] numberdata = item.Split(' ');
                number.Add(numberdata[0]);
            }

            WriteToFile("2");

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
                            pwd = splits[2],
                            group = splits[3]
                        });
                    }
                }
            }

            foreach (var item in column)
            {
                try
                {
                    ConnectionOptions options;
                    options = new ConnectionOptions();

                    options.Username = item.user;
                    options.Password = item.pwd;
                    options.EnablePrivileges = true;
                    options.Impersonation = ImpersonationLevel.Impersonate;

                    ManagementScope scope;
                    scope = new ManagementScope("\\\\" + item.ip + "\\root\\cimv2", options);
                    scope.Connect();

                    DateTime Now = DateTime.Now;

                    SelectQuery query = new SelectQuery("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3");

                    //execute the query using WMI
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                    //loop through each drive found
                    foreach (ManagementObject drive in searcher.Get())
                    {
                        driveInfo.Add(new parameter
                        {
                            ip = item.ip,
                            name = drive["name"].ToString(),
                            capacity = drive["FreeSpace"].ToString(),
                            size = Convert.ToDouble(drive["Size"])
                        });
                    }
                }
                catch (Exception ex)
                {
                    WriteToFile(ex.Message);
                    monitoringServices("DTI Check Capacity Server", "E", "Check capacity server :"+ex.Message);
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
                        capacity.Add(item2.name + "\nFree : " + Math.Round((Convert.ToDouble(item2.capacity) / 1024 / 1024 / 1024), 2) + " GB\nUsed : " + Math.Round((item2.size - Convert.ToDouble(item2.capacity)) / 1024 / 1024 / 1024, 2) + " GB \nSize : " + Math.Round((Convert.ToDouble(item2.size) / 1024 / 1024 / 1024), 2) + " GB");
                    }
                }
                msg.Add("IP : " + item.ip + " - " + item.group + "\n" + string.Join("\n", capacity) + "\n=================\n");
            }
            WriteToFile("3 "+ column.Count());

            //List<string> finaldata = new List<string>();
            //foreach (var item in driveInfo)
            //{
            //    finaldata.Add("IP : " + item.ip + " ("+item.name+") \nFree : " + Math.Round((Convert.ToDouble(item.capacity) / 1024 / 1024 / 1024),2) + " GB \nUsed : "+ Math.Round((item.size - Convert.ToDouble(item.capacity)) / 1024 / 1024 / 1024,2) + " GB \nSize : "+ Math.Round((Convert.ToDouble(item.size) / 1024 / 1024 / 1024), 2) + " GB\n");
            //}
            //SendMessage(chat_id, string.Join("\n", msg)+"\nTimeStamp : "+DateTime.Now.ToString("HH:mm"));

            //SendMessage(chat_id, "#check capacity server#\n\nPlease answer this message with format :\n*Server name\n*username\n*password");
            //SendTelegram("-342652785", string.Join("\n", msg) + "\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
            WriteData(string.Join("\n", msg));




            //Byte[] bytes = File.ReadAllBytes(AppDomain.CurrentDomain.BaseDirectory + "\\dokumen\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt");
            //String file64 = Convert.ToBase64String(bytes);
            //Console.WriteLine(file64);
            string[] splitdata = string.Join("\n", msg).Split(new string[] { "IP : 10.10.10.88 - AD01-DC" }, StringSplitOptions.None);
            
            foreach (var item in number)
            {
                //SendFileWA(item, "data:@file/plain;base64,"+file64);
                SendMessage(item, splitdata[0] + "\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
                SendMessage(item, "IP : 10.10.10.88 - AD01-DC\n"+splitdata[1] + "\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
            }
            SendTelegram("-342652785", splitdata[0] + "\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));
            SendTelegram("-342652785", "IP : 10.10.10.88 - AD01-DC\n" + splitdata[1] + "\nTimeStamp : " + DateTime.Now.ToString("HH:mm"));

            //sendFileTelegram("-342652785", AppDomain.CurrentDomain.BaseDirectory + "\\dokumen\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt");
        }
        public class data
        {
            public string ip { get; set; }
            public string user { get; set; }
            public string pwd { get; set; }
            public string group { get; set; }
        }
        public class parameter
        {
            public string ip { get; set; }
            public string name { get; set; }
            public string capacity { get; set; }
            public Double size { get; set; }
        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private static void SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac", Method.POST);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("body", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            WriteToFile(responseWa.Result.Content);

        }
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
            public int lastMessageNumber { get; set; }
        }
        public class MessageChat
        {
            public string id { get; set; }
            public string body { get; set; }
            public string fromMe { get; set; }
            public string self { get; set; }
            public string isForwarded { get; set; }
            public string author { get; set; }
            public double time { get; set; }
            public string chatId { get; set; }
            public int messageNumber { get; set; }
            public string type { get; set; }
            public string senderName { get; set; }
            public string caption { get; set; }
            public string quotedMsgBody { get; set; }
            public string quotedMsgId { get; set; }
            public string quotedMsgType { get; set; }
            public string chatName { get; set; }
        }
        private static string GetMessageList(string chatId, int lastMessageNumber)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastMessageNumber + "&chatId=" + chatId);
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.GET);
            IRestResponse responseWa = client.Execute(requestWa);
            return responseWa.Content;
        }
        protected override void OnStop()
        {
        }
        private static void SendTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;

            var client = new RestClient("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendMessage?chat_id=" + chatId + "&text=" + body);
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendMessage?chat_id=" + chatId + "&text=" + body, Method.GET);
            requestWa.Timeout = -1;
            var responseWa = client.ExecutePostAsync(requestWa);
            WriteToFile(responseWa.Result.Content);
        }
        private static string monitoringServices(string servicename, string status, string desc)
        {
            string jsonString = "{" +
                                "\"name\" : \"" + servicename + "\"," +
                                "\"logstatus\": \"" + status + "\"," +
                                "\"logdesc\":\"" + desc + "\"," +
                                "}";
            var client = new RestClient("https://apiservicekbi.azurewebsites.net/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("https://apiservicekbi.azurewebsites.net/api/ServiceStatus", Method.POST);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        public static void WriteData(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\dokumen";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\dokumen\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendDocument");
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5343027475:AAGFvuP05c_KeBVVh1Fd1hJNgWPrIH_QpjQ/sendDocument", Method.POST);


            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static void SendFileWA(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendFile?token=jkdjtwjkwq2gfkac", Method.POST);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("body", body);
            requestWa.AddParameter("filename", "ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt");
            var responseWa = client.ExecutePostAsync(requestWa);
            WriteToFile(responseWa.Result.Content);

        }
    }
}

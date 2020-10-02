using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;

namespace Logonoff
{
    public partial class Service1 : ServiceBase
    {
        [DllImport("Wtsapi32.dll")]
        private static extern bool WTSQuerySessionInformation(IntPtr hServer, int sessionId, WtsInfoClass wtsInfoClass, out IntPtr ppBuffer, out int pBytesReturned);
        [DllImport("Wtsapi32.dll")]
        private static extern void WTSFreeMemory(IntPtr pointer);

        private enum WtsInfoClass
        {
            WTSUserName = 5,
            WTSDomainName = 7,
        }

        private string GetUsername(int sessionId, bool prependDomain = true)
        {
            IntPtr buffer;
            int strLen;
            string username = "SYSTEM";
            if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WtsInfoClass.WTSUserName, out buffer, out strLen) && strLen > 1)
            {
                username = Marshal.PtrToStringAnsi(buffer);
                WTSFreeMemory(buffer);
                if (prependDomain)
                {
                    if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WtsInfoClass.WTSDomainName, out buffer, out strLen) && strLen > 1)
                    {
                        username = Marshal.PtrToStringAnsi(buffer) + "\\" + username;
                        WTSFreeMemory(buffer);
                    }
                }
            }
            return username;
        }

        private Config config;

        private void LoadConfig()
        {
            string file = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), "config.json");
            if (File.Exists(file))
            {
                string content = File.ReadAllText(file);
                try
                {
                    this.config = JsonConvert.DeserializeObject<Config>(content);
                }
                catch (Exception e)
                {
                    this.config = new Config();
                }
            }
        }

        private class Config
        {
            public Dictionary<string, string> locations { get; set; }
        }

        public Service1()
        {
            CanPauseAndContinue = true;
            CanHandleSessionChangeEvent = true;
            ServiceName = "Logonoff";
            InitializeComponent();
            this.LoadConfig();
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }

        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            string username = this.GetUsername(changeDescription.SessionId, false);

            string location = this.GetLocation();

            DateTime n = DateTime.Now;

            int num_semaine = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(n, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

            string filename = "C:\\Users\\" + username.ToLower() + "\\" + username + "-" + n.ToString("yyyy") + "-" + num_semaine.ToString() + ".csv";

            File.AppendAllText(filename, $@"{n.ToString("dd/MM/yyyy HH:mm:ss")};{changeDescription.Reason.ToString()};{location};" + "\r\n");
        }

        private string GetLocation()
        {
            string location = "";
            try
            {
                string IP = this.GetPublicIp();
                if (IP != null)
                {
                    location = IP; // A defaut
                    if (this.config.locations.ContainsKey(IP))
                    {
                        location = this.config.locations[IP];
                    }
                }
            }catch(Exception e) { }
            return location;
        }

        private string GetPublicIp()
        {
            string serviceUrl = "https://ipinfo.io/ip";
            string IP = null;
            try
            {
                IP = new System.Net.WebClient().DownloadString(serviceUrl).Trim();
            }
            catch (Exception e)
            {
            }
            return IP;
        }


    }
}

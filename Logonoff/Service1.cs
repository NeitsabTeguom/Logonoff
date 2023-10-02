using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Windows.Forms;

namespace Logonoff
{
    public partial class Service1 : ServiceBase
    {
        private string LoggedUsername;

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
            string username = null;
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


        [DllImport("kernel32.dll")]
        private static extern uint WTSGetActiveConsoleSessionId();

        private string GetCurrentUsername(int sessionId, bool prependDomain = true)
        {
            string username = null;
            username = this.GetUsername(sessionId, prependDomain);
            if (username == null || username.ToLower() == "system")
            {
                var activeSession = WTSGetActiveConsoleSessionId();
                username = this.GetUsername((int)activeSession, prependDomain);
            }
            if (username == null || username.ToLower() == "system")
            {
                username = this.config.defaultuser;
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
            public string defaultuser { get; set; }
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
            /*
            string location = this.GetLocation();

            DateTime n = DateTime.Now;

            File.AppendAllText(this.GetCurrentFile(), $@"{n.ToString("dd/MM/yyyy HH:mm:ss")};Stop;{location};" + "\r\n");
            */
        }

        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            this.LoggedUsername = this.GetCurrentUsername(changeDescription.SessionId, false);

            if (this.LoggedUsername != null)
            {
                string location = this.GetLocation();

                DateTime n = DateTime.Now;

                File.AppendAllText(this.GetCurrentFile(), $@"{n.ToString("dd/MM/yyyy HH:mm:ss")};{changeDescription.Reason.ToString()};{location};" + "\r\n");
            }
        }

        private string GetCurrentFile()
        {
            DateTime n = DateTime.Now;

            int num_semaine = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(n, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

            return Path.Combine(@"C:\Users\", this.LoggedUsername.ToLower()) + "\\" +
                this.LoggedUsername + "-" + n.ToString("yyyy") + "-" + num_semaine.ToString() + ".csv";
        }

        private string GetLocation()
        {
            string location = "";
            try
            {
                string IP = this.GetDefaultGatewayIp();
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

        private string GetDefaultGatewayIp()
        {
            IPAddress result = null;
            var cards = NetworkInterface.GetAllNetworkInterfaces().ToList();
            if (cards.Any())
            {
                foreach (var card in cards)
                {
                    var props = card.GetIPProperties();
                    if (props == null)
                        continue;

                    var gateways = props.GatewayAddresses;
                    if (!gateways.Any())
                        continue;

                    var gateway =
                        gateways.FirstOrDefault(g => g.Address.AddressFamily.ToString() == "InterNetwork");
                    if (gateway == null)
                        continue;

                    result = gateway.Address;
                    break;
                };
            }

            return result.ToString();
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

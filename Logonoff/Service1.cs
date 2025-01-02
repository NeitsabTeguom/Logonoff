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

        private class LogOnOff
        {
            public DateTime moment { get; set; }
            public string reason { get; set; }
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

            this.Save(n, "Stop", location);
            */
        }

        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            this.LoggedUsername = this.GetCurrentUsername(changeDescription.SessionId, false);

            if (this.LoggedUsername != null)
            {
                string location = this.GetLocation();

                DateTime n = DateTime.Now;

                this.Save(n, changeDescription.Reason, location);
            }
        }

        private void Save(DateTime dt, SessionChangeReason reason, string location)
        {
            string file = this.GetCurrentFile();

            /*
            string fileJSON = file + ".json";

            if (File.Exists(fileJSON))
            {
                try
                {
                    string _reason = "";

                    switch(reason)
                    {
                        case SessionChangeReason.ConsoleConnect:
                            {
                                _reason = "IN";
                                break;
                            }
                        case SessionChangeReason.ConsoleDisconnect:
                            {
                                _reason = "OUT";
                                break;
                            }
                        case SessionChangeReason.RemoteConnect:
                            {
                                _reason = "IN";
                                break;
                            }
                        case SessionChangeReason.RemoteDisconnect:
                            {
                                _reason = "OUT";
                                break;
                            }
                        case SessionChangeReason.SessionLock:
                            {
                                _reason = "OUT";
                                break;
                            }
                        case SessionChangeReason.SessionUnlock:
                            {
                                _reason = "IN";
                                break;
                            }
                        case SessionChangeReason.SessionLogon:
                            {
                                _reason = "IN";
                                break;
                            }
                        case SessionChangeReason.SessionLogoff:
                            {
                                _reason = "OUT";
                                break;
                            }
                    }

                    Dictionary<string, LogOnOff> data = JsonConvert.DeserializeObject<Dictionary<string, LogOnOff>>(File.ReadAllText(fileJSON));

                    string day = dt.ToString("dd/MM/yyyy");

                    LogOnOff loo = data.First((el) => el.Key == day && el.Value.reason == _reason).Value;
                    if(loo == null)
                    {
                        data.Add(day, new LogOnOff() { moment = dt, reason = _reason });
                    }
                    else
                    {
                        if (_reason == "OUT" && dt.TimeOfDay <= new TimeSpan(13, 0, 0))
                        {
                            loo.moment = dt;
                        }

                        if (_reason == "IN" && dt.TimeOfDay <= new TimeSpan(13, 0, 0))
                        {
                            loo.moment = dt;
                        }
                    }
                    
                    File.AppendAllText(fileJSON, JsonConvert.SerializeObject(data));
                }
                catch (Exception e)
                {
                }
            }*/

            File.AppendAllText(file + ".csv", $@"{dt.ToString("dd/MM/yyyy HH:mm:ss")};{reason.ToString()};{location};" + "\r\n");

        }

        private string GetCurrentFile()
        {

            DateTime n = DateTime.Now;

            int num_semaine = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(n, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

            string year = n.ToString("yyyy");

            if (num_semaine > 52)
            {
                year = (int.Parse(year) - 1).ToString();
            }

            return Path.Combine(@"C:\Users\", this.LoggedUsername.ToLower()) + "\\" +
                this.LoggedUsername + "-" + year + "-" + num_semaine.ToString();
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

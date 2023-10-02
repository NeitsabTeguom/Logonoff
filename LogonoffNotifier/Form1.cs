using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LogonoffNotifier
{
    public partial class Form1 : Form
    {
        // SystemTray : https://www.c-sharpcorner.com/UploadFile/f9f215/how-to-minimize-your-application-to-system-tray-in-C-Sharp/
        bool showForm = false;

        //private double DayTime = 7.5 * 60 * 60;
        private double DayTime = 8.25 * 60 * 60;
        private double TotalWeekTime = 5 * 7.5 * 60 * 60;

        public enum SessionChangeReason
        {
            //
            // Résumé :
            //     A console session has connected.
            ConsoleConnect = 1,
            //
            // Résumé :
            //     A console session has disconnected.
            ConsoleDisconnect = 2,
            //
            // Résumé :
            //     A remote session has connected.
            RemoteConnect = 3,
            //
            // Résumé :
            //     A remote session has disconnected.
            RemoteDisconnect = 4,
            //
            // Résumé :
            //     A user has logged on to a session.
            SessionLogon = 5,
            //
            // Résumé :
            //     A user has logged off from a session.
            SessionLogoff = 6,
            //
            // Résumé :
            //     A session has been locked.
            SessionLock = 7,
            //
            // Résumé :
            //     A session has been unlocked.
            SessionUnlock = 8,
            //
            // Résumé :
            //     The remote control status of a session has changed.
            SessionRemoteControl = 9,

            Stop = 10,
            Start = 11
        }
        private class Activity
        {
            public DateTime moment { get; set; }
            public SessionChangeReason reason { get; set; }
            public string location { get; set; }

            public static Activity FromCsv(string csvLine)
            {
                string[] values = csvLine.Split(';');
                Activity activity = new Activity();
                activity.moment = Convert.ToDateTime(values[0]);
                activity.reason = (SessionChangeReason)Enum.Parse(typeof(SessionChangeReason), Convert.ToString(values[1]));
                activity.location = Convert.ToString(values[2]);
                return activity;
            }
        }

        private List<Activity> activities;

        public Form1()
        {
            InitializeComponent();

            this.activities = new List<Activity>();

            this.ShowWeekHours();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ShowWeekHours();
        }

        private void ShowWeekHours()
        {
            listBox1.Items.Clear();

            this.LoadCsv();

            DateTime now = DateTime.Now;

            double WeekTime = 0;
            double CompleteWeekTime = 0;
            double TodayTime = 0;

            string currentDay = "";
            DateTime? dayBegin = null;
            DateTime? dayPause = null;
            DateTime? dayResume = null;
            DateTime? dayEnd = null;

            List<Activity> activities = this.activities.OrderBy(a => a.moment).ToList();
            DateTime? startDT = null;
            foreach (Activity activity in activities)
            {
                // Compteur

                if (startDT == null && (activity.reason == SessionChangeReason.SessionLogon || activity.reason == SessionChangeReason.SessionUnlock))
                {
                    startDT = activity.moment;
                }

                if (startDT != null && (activity.reason == SessionChangeReason.SessionLogoff || activity.reason == SessionChangeReason.SessionLock))
                {
                    if (startDT != null)
                    {
                        WeekTime += activity.moment.Subtract((DateTime)startDT).TotalSeconds;

                        if (((DateTime)startDT).ToString("dd/MM/yyyy") != now.ToString("dd/MM/yyyy"))
                        {
                            CompleteWeekTime += activity.moment.Subtract((DateTime)startDT).TotalSeconds;
                        }

                        if (((DateTime)startDT).ToString("dd/MM/yyyy") == now.ToString("dd/MM/yyyy"))
                        {
                            TodayTime += activity.moment.Subtract((DateTime)startDT).TotalSeconds;
                        }

                        startDT = null;
                    }
                }

                // Résumé semaine
                if (currentDay != activity.moment.ToString("dd/MM/yyyy"))
                {
                    // Compte du jour
                    if(currentDay != "")
                    {
                        if (dayBegin != null && dayPause != null && dayResume != null && dayEnd != null)
                        {
                            double dayCount = ((DateTime)dayPause).Subtract(((DateTime)dayBegin)).TotalSeconds + ((DateTime)dayEnd).Subtract(((DateTime)dayResume)).TotalSeconds;
                            listBox1.Items.Add(currentDay + " : " + TotalSecondsToString(dayCount) + " => " +
                                TotalSecondsToString(((DateTime)dayBegin).TimeOfDay.TotalSeconds) + " - " +
                                TotalSecondsToString(((DateTime)dayPause).TimeOfDay.TotalSeconds) + " / " +
                                TotalSecondsToString(((DateTime)dayResume).TimeOfDay.TotalSeconds) + " - " +
                                TotalSecondsToString(((DateTime)dayEnd).TimeOfDay.TotalSeconds));
                        }
                    }

                    currentDay = activity.moment.ToString("dd/MM/yyyy");

                    dayBegin = activity.moment;
                    dayPause = null;
                    dayResume = null;
                    dayEnd = null;
                }
                else
                {
                    if (dayPause == null && dayResume == null && activity.moment.TimeOfDay >= new TimeSpan(11, 45, 00) && (activity.reason == SessionChangeReason.SessionLogoff || activity.reason == SessionChangeReason.SessionLock))
                    {
                        dayPause = activity.moment;
                    }

                    if (dayResume == null && activity.moment.TimeOfDay >= new TimeSpan(12, 30, 00) && (activity.reason == SessionChangeReason.SessionLogon || activity.reason == SessionChangeReason.SessionUnlock))
                    {
                        dayResume = activity.moment;
                    }

                    if (dayResume != null && (activity.reason == SessionChangeReason.SessionLogoff || activity.reason == SessionChangeReason.SessionLock || activity.reason == SessionChangeReason.ConsoleDisconnect))
                    {
                        dayEnd = activity.moment;
                    }
                }
            }

            if (dayBegin != null && dayPause != null && dayResume != null && dayEnd != null)
            {
                double dayCount = ((DateTime)dayPause).Subtract(((DateTime)dayBegin)).TotalSeconds + ((DateTime)dayEnd).Subtract(((DateTime)dayResume)).TotalSeconds;
                listBox1.Items.Add(currentDay + " : " + TotalSecondsToString(dayCount) + " => " +
                    TotalSecondsToString(((DateTime)dayBegin).TimeOfDay.TotalSeconds) + " - " +
                    TotalSecondsToString(((DateTime)dayPause).TimeOfDay.TotalSeconds) + " / " +
                    TotalSecondsToString(((DateTime)dayResume).TimeOfDay.TotalSeconds) + " - " +
                    TotalSecondsToString(((DateTime)dayEnd).TimeOfDay.TotalSeconds));
            }

            if (startDT != null && ((DateTime)startDT).ToString("dd/MM/yyyy") == now.ToString("dd/MM/yyyy"))
            {
                WeekTime += now.Subtract((DateTime)startDT).TotalSeconds;
                TodayTime += now.Subtract((DateTime)startDT).TotalSeconds;
                startDT = null;
            }

            this.label1.Text = $@"Vous avez travaillé {TotalSecondsToString(WeekTime)} cette semaine !";

            this.label2.Text = $@"Vous avez travaillé {TotalSecondsToString(TodayTime)} aujourd'hui !";


            // Calcul de l'heure de sortie pour finir correctement le nombre d'heures de la semaine

            int dayOfWeek = (((int)now.DayOfWeek + 6) % 7) + 1;

            double dayOfWeekTime = dayOfWeek * DayTime;

            double toDoToday = dayOfWeekTime - WeekTime;

            this.label5.Text = "";

            if (toDoToday > 0)
            {
                this.label3.Text = $@"Reste à travailler {TotalSecondsToString(toDoToday)} aujourd'hui pour respecter le taux horaire hebdomadaire !";
                this.label4.Text = $@"Il faut finir à {now.AddSeconds(toDoToday).ToString(@"HH:mm:ss")} aujourd'hui pour respecter le taux horaire hebdomadaire !";

                this.label5.Text = $@"{TotalSecondsToString(CompleteWeekTime - ((dayOfWeek - 1) * DayTime))} supplémentaire estimé cette semaine.";
            }
            else
            {
                this.label3.Text = $@"{TotalSecondsToString(toDoToday * -1)} supplémentaires cette semaine !";
                this.label4.Text = $@"Vous auriez du finir à {now.AddSeconds(toDoToday).ToString(@"HH:mm:ss")} aujourd'hui pour respecter le taux horaire hebdomadaire !";

                this.label5.Text = $@"{TotalSecondsToString(toDoToday * -1)} supplémentaire cette semaine.";
            }

        }

        private void LoadCsv()
        {
            this.activities = File.ReadAllLines(this.GetCurrentFile())
                                            .Select(v => Activity.FromCsv(v))
                                            .ToList();
        }
        private string GetCurrentFile()
        {
            //Environment.UserDomainName + "\\" + Environment.UserName

            DateTime n = DateTime.Now;

            int num_semaine = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(n, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

            return Path.Combine(Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)), Environment.UserName.ToLower()) + "\\" +
                Environment.UserName + "-" + n.ToString("yyyy") + "-" + num_semaine.ToString() + ".csv";
        }

        private string TotalSecondsToString(double totalSecondes)
        {
            double hours = Math.Truncate(totalSecondes / (60 * 60));
            double minutes = Math.Truncate((totalSecondes - (hours * 60 * 60)) / 60);
            double secondes = Math.Truncate(totalSecondes - (hours * 60 * 60) - (minutes * 60));

            return hours.ToString() + ":" + minutes.ToString().PadLeft(2, '0') + ":" + secondes.ToString().PadLeft(2, '0');
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            showForm = false;
            this.Visible = showForm;
            e.Cancel = true;
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            showForm = true;
            this.Visible = showForm;

            ShowWeekHours();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(showForm);
        }
    }
}

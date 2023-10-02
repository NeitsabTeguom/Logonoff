using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Logonoff
{
    class Notifier
    {
        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern bool WTSSendMessage(
           IntPtr hServer,
           [MarshalAs(UnmanagedType.I4)] int SessionId,
           String pTitle,
           [MarshalAs(UnmanagedType.U4)] int TitleLength,
           String pMessage,
           [MarshalAs(UnmanagedType.U4)] int MessageLength,
           [MarshalAs(UnmanagedType.U4)] int Style,
           [MarshalAs(UnmanagedType.U4)] int Timeout,
           [MarshalAs(UnmanagedType.U4)] out int pResponse,
           bool bWait);
        public static IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;
        public static int WTS_CURRENT_SESSION = 1;

        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern int WTSEnumerateSessions(
                System.IntPtr hServer,
                int Reserved,
                int Version,
                ref System.IntPtr ppSessionInfo,
                ref int pCount);

        [DllImport("wtsapi32.dll")]
        static extern void WTSFreeMemory(IntPtr pMemory);

        [StructLayout(LayoutKind.Sequential)]
        private struct WTS_SESSION_INFO
        {
            public Int32 SessionID;

            [MarshalAs(UnmanagedType.LPStr)]
            public String pWinStationName;

            public WTS_CONNECTSTATE_CLASS State;
        }

        public enum WTS_CONNECTSTATE_CLASS
        {
            WTSActive,
            WTSConnected,
            WTSConnectQuery,
            WTSShadow,
            WTSDisconnected,
            WTSIdle,
            WTSListen,
            WTSReset,
            WTSDown,
            WTSInit
        }

        public static List<Int32> GetActiveSessions(IntPtr server)
        {
            List<Int32> ret = new List<int>();

            try
            {
                IntPtr ppSessionInfo = IntPtr.Zero;

                Int32 count = 0;
                Int32 retval = WTSEnumerateSessions(server, 0, 1, ref ppSessionInfo, ref count);
                Int32 dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));

                Int64 current = (int)ppSessionInfo;

                if (retval != 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure((System.IntPtr)current, typeof(WTS_SESSION_INFO));
                        current += dataSize;

                        if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive)
                            ret.Add(si.SessionID);
                    }

                    WTSFreeMemory(ppSessionInfo);
                }
            }
            catch { }

            return ret;
        }

        public void Notify(string title, string message)
        {
            List<Int32> sessions = GetActiveSessions(WTS_CURRENT_SERVER_HANDLE);
            if (sessions.Count > 0)
            {
                foreach (Int32 sessionId in sessions)
                {
                        try
                        {
                            bool result = false;
                            int tlen = title.Length;
                            int mlen = message.Length;
                            //int resp = 7;
                            int resp = 0;
                            result = WTSSendMessage(WTS_CURRENT_SERVER_HANDLE, sessionId, title, tlen, message, mlen, 4, 0, out resp, true);
                            int err = Marshal.GetLastWin32Error();
                            if (err == 0)
                            {
                                if (result) //user responded to box
                                {
                                    if (resp == 7) //user clicked no
                                    {

                                    }
                                    else if (resp == 6) //user clicked yes
                                    {

                                    }
                                    //Debug.WriteLine("user_session:" + user_session + " err:" + err + " resp:" + resp);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //Debug.WriteLine("no such thread exists", ex);
                        }
                        //Application App = new Application();
                        //App.Run(new MessageForm());
                }
            }
        }
    }
}

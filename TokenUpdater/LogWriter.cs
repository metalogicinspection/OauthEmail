using System;
using System.Diagnostics;
using System.IO;

namespace TokenUpdater
{
    internal class LogWriter
    {
        private static int _eventId = 1;

       
        private const string PreferedEventSourceName = "Metalogic";

        public static void WriteLog(Exception ex, EventLogEntryType eventype = EventLogEntryType.Error)
        {
            WriteLog(string.Concat(ex.Message, Environment.NewLine, ex.StackTrace), eventype);
        }

        public static void WriteLog(string message, Exception ex, EventLogEntryType eventype = EventLogEntryType.Error)
        {
            WriteLog(string.Concat(message, Environment.NewLine, ex.Message, Environment.NewLine, ex.StackTrace), eventype);
        }

        private bool _firtTimeLog = true;

        public static void WriteLog(string message, EventLogEntryType eventype = EventLogEntryType.Information)
        {
            try
            {
                using (var eventLog = new EventLog("Application"))
                {
                    eventLog.Source = PreferedEventSourceName;
                    eventLog.WriteEntry(message, eventype, ++_eventId, 1);
                }
            }
            catch (Exception e)
            {
                using (var eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(message, eventype, ++_eventId, 1);
                }
            }
            Console.WriteLine(message);
        }
    }
}

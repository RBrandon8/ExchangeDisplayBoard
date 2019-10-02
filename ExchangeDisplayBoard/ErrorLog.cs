using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Diagnostics;

namespace ExchangeDisplayBoard
{
    public static class ErrorLog
    {
        private static readonly string source = "DisplayBoard";
        private static readonly string log = "Application";

        public static void WriteInformational(string description)
        {
            if (!EventLog.SourceExists(source))
                EventLog.CreateEventSource(source, log);
            EventLog.WriteEntry(source, description);
        }

        public static void WriteError(string description )
        {
            if (!EventLog.SourceExists(source))
                EventLog.CreateEventSource(source, log);
            EventLog.WriteEntry(source, description, EventLogEntryType.Error);
        }


    }
}
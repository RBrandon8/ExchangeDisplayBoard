using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;

namespace ExchangeDisplayBoard
{
    public class Appt : IEquatable<Appt>
    {
        public DateTime startTime { get; set; }
        public DateTime endTime { get; set; }
        public string dateTime { get; set; }
        public string description { get; set; }
        public string resource { get; set; }

        public Appt(string res, DateTime startTime, DateTime endTime, string desc)
        {
            resource = res;
            this.startTime = startTime;
            this.endTime = endTime;
            dateTime = startTime.ToShortTimeString() + '-' + endTime.ToShortTimeString();
            description = desc;
        }

        public bool Equals(Appt other)
        {
            return dateTime == other.dateTime && description == other.description && resource == other.resource;
        }

    }

}
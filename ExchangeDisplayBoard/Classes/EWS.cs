using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeDisplayBoard
{   
    public static class EWS
    {
        public static EmailAddressCollection roomsList { get; set; }
        public static BindingList<EmailAddress> rooms { get; set; }
        public static BindingList<Appt> displayApps { get; set; }
        public static readonly List<EmailAddress> privateRooms = new List<EmailAddress> {}; //Add Private Room email address that shouldn't be shown

        /* Saves available roomlists from Exchange to roomsList collection */
        public static void GetExchangeRoomList(ExchangeService service)
        {
            roomsList = service.GetRoomLists();
        }

        /* Saves a list of rooms that are not inn the privateRooms collectionn for a roomsList */
        public static void GetExchangeRooms(ExchangeService service, EmailAddress roomListAddress)
        {
            var roomsEmail = service.GetRooms(roomListAddress.Address);
            rooms = new BindingList<EmailAddress>(roomsEmail);

            //Remove Rooms that should not show on public board
            foreach (EmailAddress email in privateRooms)
            {
                var roomAdd = rooms.Where(x => x.Address == email.Address).FirstOrDefault();
                if (roomAdd != null)
                    rooms.Remove(roomAdd);
            }
        }

        /* Populates the appointments for an Exchange room in the Apps class */
        public static void GetResourceCalendarItems(ExchangeService service, string room, DateTime startDate, DateTime endDate)
        {
            try
            {
                CalendarView cv = new CalendarView(startDate, endDate);
                String MailboxToAccess = room;
                FolderId CalendarFolderId = new FolderId(WellKnownFolderName.Calendar, MailboxToAccess);
                CalendarFolder calendar = CalendarFolder.Bind(service, CalendarFolderId);
                FindItemsResults<Appointment> fapts = calendar.FindAppointments(cv);
                if (fapts.Items.Count > 0)
                {
                    foreach (Appointment a in fapts)
                    {
                        Appt appointment = new Appt(a.Location, a.Start, a.End, a.Subject.ToString());

                        if (!a.Subject.ToUpper().Contains("SET-UP") && !a.Subject.ToUpper().Contains("SETUP") && !a.Subject.ToUpper().Contains("SET UP")
                            && !a.Subject.ToUpper().Contains("TEARDOWN") && !a.Subject.ToUpper().Contains("TEAR-DOWN") && !a.Subject.ToUpper().Contains("TEAR DOWN"))
                        {
                            if (!displayApps.Contains(appointment))
                                displayApps.Add(appointment);
                        }
                    }
                }
            }
            catch (Exception e) { }
        }

        public static void SortAppsList()
        {
            List<Appt> listApps = displayApps.ToList();
            listApps.Sort((Appt X, Appt Y) => X.startTime.CompareTo(Y.startTime));
            displayApps = new BindingList<Appt>(listApps);

        }
    }








}
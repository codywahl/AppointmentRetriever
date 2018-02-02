using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AppointmentRetriever.Services
{
    public class AppointmentManager
    {
        #region Fields

        private readonly List<Appointment> _appointments;

        #endregion Fields

        #region Constructors

        public AppointmentManager()
        {
            _appointments = new List<Appointment>();
        }

        public AppointmentManager(List<Appointment> appointments)
        {
            _appointments = appointments;
        }

        #endregion Constructors

        public void AddOrUpdateAppointment(Appointment exchangeVersionOfAppointment)
        {
            var ourVersionOfAppointment = GetAppointmentWithICalUid(exchangeVersionOfAppointment.ICalUid);

            if (ourVersionOfAppointment == null)
            {
                // Create a private Add method which contains logic to send room access message
                _appointments.Add(exchangeVersionOfAppointment);
                return;
            }

            if (!WeHaveTheCurrentVersion(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                // Determine differences
                // Act on differences
                // when done, update our version of appointment

                SendAppointmentDeltaEvents(ourVersionOfAppointment, exchangeVersionOfAppointment);

                UpdateAppointment(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }
        }

        public void RemoveAppointment(Appointment appointment)
        {
            var index = _appointments.IndexOf(appointment);
            if (index != -1)
            {
                // Create a private Remove method which contains logic to send room access removal message
                _appointments.RemoveAt(index);
            }
        }

        public int GetAppointmentCount()
        {
            return _appointments.Count;
        }

        #region Private Helpers

        private Appointment GetAppointmentWithICalUid(string iCalUid)
        {
            return _appointments.FirstOrDefault(x => x.ICalUid.Equals(iCalUid));
        }

        //private bool ContainsAppointment(Appointment exchangeVersionOfAppointment)
        //{
        //    if (exchangeVersionOfAppointment == null)
        //    {
        //        throw new ArgumentNullException(nameof(exchangeVersionOfAppointment));
        //    }

        //    return GetAppointmentWithId(exchangeVersionOfAppointment.Id.UniqueId) != null;
        //}

        private bool WeHaveTheCurrentVersion(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            if (ourVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(ourVersionOfAppointment));
            }
            if (exchangeVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(exchangeVersionOfAppointment));
            }

            return ourVersionOfAppointment.Id.SameIdAndChangeKey(exchangeVersionOfAppointment.Id);
        }

        private void UpdateAppointment(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            if (ourVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(ourVersionOfAppointment));
            }
            if (exchangeVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(exchangeVersionOfAppointment));
            }

            Console.WriteLine("\t(AppointmentManager:Updating exchangeVersionOfAppointment...)");

            var index = _appointments.IndexOf(ourVersionOfAppointment);
            if (index != -1)
            {
                _appointments[index] = exchangeVersionOfAppointment;
            }
        }

        private static void SendAppointmentDeltaEvents(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            if (MeetingRoomChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendMeetingRoomChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }

            if (MeetingDateChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendMeetingDateChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }

            if (AttendeesChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAttendeesChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }

            if (AttendeeResponseChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAttendeeResponseChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }
        }

        private static bool MeetingRoomChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            return !ourVersionOfAppointment.Location.Equals(exchangeVersionOfAppointment.Location);
        }

        private static void SendMeetingRoomChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            Console.WriteLine(">> Meeting Room Changed Event Sent...");
        }

        private static bool MeetingDateChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            var startChanged = !ourVersionOfAppointment.Start.Date.Equals(exchangeVersionOfAppointment.Start.Date);
            var endChanged = !ourVersionOfAppointment.End.Date.Equals(exchangeVersionOfAppointment.End.Date);

            return startChanged || endChanged;
        }

        private static void SendMeetingDateChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            Console.WriteLine(">> Meeting Date Changed Event Sent...");
        }

        private static bool AttendeesChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            var requireAttendeesChanged = !ourVersionOfAppointment.RequiredAttendees.Count.Equals(exchangeVersionOfAppointment.RequiredAttendees.Count());
            var optionalAttendeesChanged = !ourVersionOfAppointment.OptionalAttendees.Count.Equals(exchangeVersionOfAppointment.OptionalAttendees.Count());

            return requireAttendeesChanged || optionalAttendeesChanged;
        }

        private static void SendAttendeesChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            Console.WriteLine(">> Meeting Attendees Changed Event Sent...");
        }

        private static bool AttendeeResponseChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            // WORK IN PROGRESS

            //var changedResponseType = false;

            //foreach (var ourAttendee in ourVersionOfAppointment.RequiredAttendees)
            //{
            //    var exchangeAttendeeIndex = exchangeVersionOfAppointment.RequiredAttendees.IndexOf(ourAttendee);

            //    if (exchangeAttendeeIndex != -1)
            //    {
            //        var exchangeAttendee = exchangeVersionOfAppointment.RequiredAttendees[exchangeAttendeeIndex];

            //        if (ourAttendee.ResponseType.HasValue)
            //        {
            //        }
            //    }

            //    if (changedResponseType)
            //    {
            //        return true;
            //    }
            //}

            return false;
        }

        private static void SendAttendeeResponseChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            Console.WriteLine(">> Meeting Attendee(s) Response Changed Event Sent...");
        }

        #endregion Private Helpers
    }
}
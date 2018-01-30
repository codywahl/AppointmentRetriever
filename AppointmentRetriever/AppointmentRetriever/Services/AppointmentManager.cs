using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AppointmentRetriever.Services
{
    public class AppointmentManager
    {
        private readonly List<Appointment> _appointments;

        public AppointmentManager()
        {
            _appointments = new List<Appointment>();
        }

        public AppointmentManager(List<Appointment> appointments)
        {
            _appointments = appointments;
        }

        public void AddOrUpdateAppointment(Appointment appointment)
        {
            var ourVersionOfAppointment = GetAppointmentWithICalUid(appointment.ICalUid);

            if (ourVersionOfAppointment == null)
            {
                Console.WriteLine("\t(AppointmentManager: Adding appointment...)");
                _appointments.Add(appointment);
                return;
            }

            if (!WeHaveTheCurrentVersion(ourVersionOfAppointment, appointment))
            {
                Console.WriteLine("((Appointment change detected. Send message to bus or whatever.))");
                UpdateAppointment(ourVersionOfAppointment, appointment);
            }
        }

        public void RemoveAppointment(Appointment appointment)
        {
            var index = _appointments.IndexOf(appointment);
            if (index != -1)
            {
                _appointments.RemoveAt(index);
            }
        }

        public int GetAppointmentCount()
        {
            return _appointments.Count;
        }

        private Appointment GetAppointmentWithICalUid(string iCalUid)
        {
            return _appointments.FirstOrDefault(x => x.ICalUid.Equals(iCalUid));
        }

        //private bool ContainsAppointment(Appointment appointment)
        //{
        //    if (appointment == null)
        //    {
        //        throw new ArgumentNullException(nameof(appointment));
        //    }

        //    return GetAppointmentWithId(appointment.Id.UniqueId) != null;
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

            Console.WriteLine("\t(AppointmentManager:Updating appointment...)");

            var index = _appointments.IndexOf(ourVersionOfAppointment);
            if (index != -1)
            {
                _appointments[index] = exchangeVersionOfAppointment;
            }
        }
    }
}
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AppointmentRetriever.Services
{
    public class AppointmentManager
    {
        #region Fields

        // This list of appointments should be kept in RC and retrieved by this application periodically.
        // Doing this will keep us in sync.
        private static readonly List<Appointment> Appointments = new List<Appointment>();

        #endregion Fields

        /// <summary>
        /// Adds the provided appointment to our stored list if not found.
        /// If found, we check the version of our appointment, updating if not the same.
        /// </summary>
        /// <param name="exchangeVersionOfAppointment">The appointment to be updated or added.</param>
        public void AddOrUpdateAppointment(Appointment exchangeVersionOfAppointment)
        {
            var ourVersionOfAppointment = GetOurVersionOfAppointment(exchangeVersionOfAppointment);

            if (ourVersionOfAppointment == null)
            {
                AddAppointment(exchangeVersionOfAppointment);
                return;
            }

            if (!WeHaveTheCurrentVersion(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAppointmentDeltaEvents(ourVersionOfAppointment, exchangeVersionOfAppointment);
                UpdateAppointment(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }
        }

        /// <summary>
        /// Adds or updates the provided appointments to our stored list if not found.
        /// </summary>
        /// <param name="exchangeVersionOfAppointments">The appointments to be updated or added.</param>
        public void AddOrUpdateAppointments(List<Appointment> exchangeVersionOfAppointments)
        {
            CheckForDeletedOrPastAppointments(exchangeVersionOfAppointments);

            foreach (var exchangeVersionOfAppointment in exchangeVersionOfAppointments)
            {
                AddOrUpdateAppointment(exchangeVersionOfAppointment);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="exchangeAppointments"></param>
        private static void CheckForDeletedOrPastAppointments(IReadOnlyCollection<Appointment> exchangeAppointments)
        {
            var deletedOrPastAppointments = Appointments.Where(x => exchangeAppointments.All(y => !IsTheSameAppointment(x, y))).ToList();

            foreach (var deletedOrPastAppointment in deletedOrPastAppointments)
            {
                if (deletedOrPastAppointment.End.Date >= DateTime.Today.Date)
                {
                    // If the appointment we have is in the future, but we didn't find it in exchange, pressume it was deleted.
                    SendAppointmentDeletedEvent(deletedOrPastAppointment);
                }

                RemoveAppointment(deletedOrPastAppointment);
            }
        }

        /// <summary>
        /// Gets the number of appointments we have.
        /// </summary>
        /// <returns>Appointment count</returns>
        public int GetAppointmentCount()
        {
            return Appointments.Count;
        }

        public string AppointmentsToString()
        {
            try
            {
                var output = $"{Environment.NewLine}\tAppointments{Environment.NewLine}";

                foreach (var appointment in Appointments)
                {
                    output += $"\t-----------------------------------------------------------------{Environment.NewLine}";
                    output += $"\t           Subject: {appointment.Subject}{Environment.NewLine}";
                    output += $"\t        Start Time: {appointment.Start}{Environment.NewLine}";
                    output += $"\t          End Time: {appointment.End}{Environment.NewLine}";
                    output += $"\t          Location: {appointment.Location}{Environment.NewLine}";
                    output += $"\t           ICal Id: ...{appointment.ICalUid.Substring(appointment.ICalUid.Length - 30)}{Environment.NewLine}";
                    output += $"\t         Unique Id: ...{appointment.Id.UniqueId.Substring(appointment.Id.UniqueId.Length - 30)}{Environment.NewLine}";
                    output += $"\t         Change Id: {appointment.Id.ChangeKey}{Environment.NewLine}";
                    output += $"\t         Organizer: {appointment.Organizer.Name} {appointment.Organizer.Address}{Environment.NewLine}";
                    output += $"\tRequired Attendees: {GetAttendeeNamesAndResponse(appointment.RequiredAttendees)}{Environment.NewLine}";
                    output += $"\tOptional Attendees: {GetAttendeeNamesAndResponse(appointment.OptionalAttendees)}{Environment.NewLine}";
                    output += $"\t-----------------------------------------------------------------{Environment.NewLine}";
                }

                return output;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine();
                return null;
            }
        }

        #region Private Helpers

        #region Add, Update, Remove

        /// <summary>
        /// Adds an appointment.
        /// </summary>
        /// <param name="appointmentToAdd"></param>
        private static void AddAppointment(Appointment appointmentToAdd)
        {
            // Instead of adding to this list, we would refresh from RC after the event was sent
            Appointments.Add(appointmentToAdd);

            SendNewAppointmentEvent(appointmentToAdd);
        }

        /// <summary>
        /// Updates an appointment.
        /// </summary>
        /// <param name="ourVersionOfAppointment"></param>
        /// <param name="exchangeVersionOfAppointment"></param>
        private static void UpdateAppointment(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            Console.WriteLine("\t(AppointmentManager:Updating exchangeVersionOfAppointment...)");

            var index = Appointments.IndexOf(ourVersionOfAppointment);
            if (index != -1)
            {
                Appointments[index] = exchangeVersionOfAppointment;
            }
        }

        /// <summary>
        /// Removes an appointment.
        /// </summary>
        /// <param name="appointmentToRemove"></param>
        private static void RemoveAppointment(Appointment appointmentToRemove)
        {
            var index = Appointments.IndexOf(appointmentToRemove);
            if (index != -1)
            {
                Appointments.RemoveAt(index);
            }
        }

        #endregion Add, Update, Remove

        /// <summary>
        /// Determines is two appointments are the same by comparing their ICalUId and ICalReccurenceId.
        /// </summary>
        /// <param name="appointment1"></param>
        /// <param name="appointment2"></param>
        /// <returns></returns>
        private static bool IsTheSameAppointment(Appointment appointment1, Appointment appointment2)
        {
            return appointment1.ICalUid.Equals(appointment2.ICalUid) &&
                   appointment1.ICalRecurrenceId.Equals(appointment2.ICalRecurrenceId);
        }

        /// <summary>
        /// Gets our version of an appointment.
        /// Note: Our version may not be the same version of the appointment we are looking for in our stored appointments.
        /// </summary>
        /// <param name="exchangeVersionOfAppointment"></param>
        /// <returns>Our version of the given appointment or null if not found.</returns>
        private static Appointment GetOurVersionOfAppointment(Appointment exchangeVersionOfAppointment)
        {
            return Appointments.FirstOrDefault(x => IsTheSameAppointment(exchangeVersionOfAppointment, x));
        }

        /// <summary>
        /// Checks if the two appointments are the same version or not.
        /// </summary>
        /// <param name="appointment1"></param>
        /// <param name="appointment2"></param>
        /// <returns></returns>
        private bool WeHaveTheCurrentVersion(Appointment appointment1, Appointment appointment2)
        {
            if (appointment1 == null)
            {
                throw new ArgumentNullException(nameof(appointment1));
            }
            if (appointment2 == null)
            {
                throw new ArgumentNullException(nameof(appointment2));
            }
            if (!IsTheSameAppointment(appointment1, appointment2))
            {
                throw new ArgumentException("The two appointments were not the same and should be when comparing versions.");
            }

            return appointment1.Id.SameIdAndChangeKey(appointment2.Id);
        }

        private static void SendAppointmentDeltaEvents(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            if (AppointmentIsCancelled(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAppointmentCancelledEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
                RemoveAppointment(ourVersionOfAppointment);
                return;
            }

            if (AppointmentLocationOrDateChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAppointmentLocationOrDateChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
                return;
            }

            if (AppointmentAttendeesChanged(ourVersionOfAppointment, exchangeVersionOfAppointment))
            {
                SendAppointmentAttendeesChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment);
            }

            // We determined that only an organizer's mailbox contains the responses from each attendee. Because of this, we may not do this.
            //var attendeesWithUpdatedResponse = GetAppointmentAttendeesWithResponseChanged(ourVersionOfAppointment, exchangeVersionOfAppointment);
            //if (attendeesWithUpdatedResponse.Any())
            //{
            //    SendAppointmentAttendeeResponseChangedEvent(ourVersionOfAppointment, exchangeVersionOfAppointment, attendeesWithUpdatedResponse);
            //}
        }

        private static void SendNewAppointmentEvent(Appointment newAppointment)
        {
            if (newAppointment == null)
            {
                throw new ArgumentNullException(nameof(newAppointment));
            }

            Console.WriteLine(">> New Appointment Event Sent...");
        }

        #region Date & Location

        private static bool AppointmentLocationOrDateChanged(Appointment ourVersionOfAppointment,
            Appointment exchangeVersionOfAppointment)
        {
            return AppointmentLocationChanged(ourVersionOfAppointment, exchangeVersionOfAppointment) ||
                   AppointmentDateChanged(ourVersionOfAppointment, exchangeVersionOfAppointment);
        }

        private static bool AppointmentLocationChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            return !ourVersionOfAppointment.Location.Equals(exchangeVersionOfAppointment.Location);
        }

        private static bool AppointmentDateChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            var startChanged = !ourVersionOfAppointment.Start.Date.Equals(exchangeVersionOfAppointment.Start.Date);
            var endChanged = !ourVersionOfAppointment.End.Date.Equals(exchangeVersionOfAppointment.End.Date);

            return startChanged || endChanged;
        }

        private static void SendAppointmentLocationOrDateChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            Console.WriteLine(">> Appointment Location Or Date Changed Event Sent...");
        }

        #endregion Date & Location

        #region Attendees

        private static bool AppointmentAttendeesChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            var requireAttendeesChanged =
                !ourVersionOfAppointment.RequiredAttendees.Count.Equals(exchangeVersionOfAppointment.RequiredAttendees
                    .Count());
            var optionalAttendeesChanged =
                !ourVersionOfAppointment.OptionalAttendees.Count.Equals(exchangeVersionOfAppointment.OptionalAttendees
                    .Count());

            return requireAttendeesChanged || optionalAttendeesChanged;
        }

        private static void SendAppointmentAttendeesChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            Console.WriteLine(">> Appointment Attendees Changed Event Sent...");
        }

        private static List<Attendee> GetAppointmentAttendeesWithResponseChanged(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            var attendeesWithUpdatedResponse = exchangeVersionOfAppointment.RequiredAttendees.Where(x =>
                ourVersionOfAppointment.RequiredAttendees.All(y => y.Address.Equals(x.Address) && y.ResponseType != x.ResponseType)).ToList();

            attendeesWithUpdatedResponse.AddRange(exchangeVersionOfAppointment.OptionalAttendees.Where(x =>
                ourVersionOfAppointment.OptionalAttendees.All(y => y.Address.Equals(x.Address) && y.ResponseType != x.ResponseType)).ToList());

            return attendeesWithUpdatedResponse;
        }

        //private static void SendAppointmentAttendeeResponseChangedEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment, List<Attendee> attendeesWithUpdatedResponse)
        //{
        //    AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

        //    Console.WriteLine(">> Appointment Attendee(s) Response Changed Event Sent...");
        //}

        #endregion Attendees

        #region Canceled/Deleted

        private static bool AppointmentIsCancelled(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            return exchangeVersionOfAppointment.IsCancelled;
        }

        private static void SendAppointmentCancelledEvent(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            AppointmentsEventPreconditionCheck(ourVersionOfAppointment, exchangeVersionOfAppointment);

            Console.WriteLine(">> Appointment Cancelled Event Sent...");
        }

        private static void SendAppointmentDeletedEvent(Appointment deletedAppointment)
        {
            if (deletedAppointment == null)
            {
                throw new ArgumentNullException(nameof(deletedAppointment));
            }

            Console.WriteLine(">> Appointment Deleted Event Sent...");
        }

        #endregion Canceled/Deleted

        private static void AppointmentsEventPreconditionCheck(Appointment ourVersionOfAppointment, Appointment exchangeVersionOfAppointment)
        {
            if (ourVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(ourVersionOfAppointment));
            }
            if (exchangeVersionOfAppointment == null)
            {
                throw new ArgumentNullException(nameof(exchangeVersionOfAppointment));
            }
        }

        private static string GetAttendeeNamesAndResponse(AttendeeCollection attendees)
        {
            var nameList = new List<string>();
            foreach (var attendee in attendees)
            {
                var s = attendee.Name + " -> ";

                if (attendee.ResponseType.HasValue)
                {
                    s += attendee.ResponseType.Value.ToString();
                }
                else
                {
                    s += "None";
                }

                nameList.Add(s);
            }

            return string.Join(",", nameList);
        }

        #endregion Private Helpers
    }
}
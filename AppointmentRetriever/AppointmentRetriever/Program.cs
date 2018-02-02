using AppointmentRetriever.Services;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace AppointmentRetriever
{
    public class Program
    {
        private static AppointmentManager _appointmentManager;
        private static CancellationTokenSource _cancellationTokenSource;

        private static void Main(string[] args)
        {
            _appointmentManager = new AppointmentManager();
            _cancellationTokenSource = new CancellationTokenSource();

            var getMeetingsTask = new Task(async () => await GetMeetingsLoop(), _cancellationTokenSource.Token);
            var uiLoop = new Task(async () => await UiLoop());

            getMeetingsTask.Start();
            uiLoop.Start();

            Task.WaitAll(uiLoop, getMeetingsTask);

            _cancellationTokenSource.Dispose();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        #region Private Helpers

        private static void CancelMeeting()
        {
            MeetingRoomsService.CancelAppointment(@"AQMkAGMzZjAzMTI4LTIzODQtNDdhMi1hNTFlLWUzOGE0NTk2NTlmMABGAAADrvpyiRwtfEOF5Q/Zf/He0wcARDWNuRjujk23tS/kvy7RNwAAAw8AAABENY25GO6OTbe1L+S/LtE3AAADMAAAAA==", "Cuz i want to.");
        }

        private static async Task UiLoop()
        {
            Console.WriteLine(">> Starting UI Loop ");

            while (!_cancellationTokenSource.IsCancellationRequested)
            {
                var input = Console.ReadLine();

                if (input != null && input.ToLower().Equals("quit"))
                {
                    _cancellationTokenSource.Cancel();
                }
            }

            await Task.Delay(1);
        }

        private static async Task GetMeetingsLoop()
        {
            Console.WriteLine(">> Starting GetMeetings Loop ");

            while (!_cancellationTokenSource.IsCancellationRequested)
            {
                GetMeetings();
            }

            await Task.Delay(5000);
        }

        private static void GetMeetings()
        {
            var output = $"{Environment.NewLine}{"Room",-35}{"Appointment Count",-45}{Environment.NewLine}";
            output += $"-----------------------------------------------------------------------{Environment.NewLine}{Environment.NewLine}";

            var rooms = MeetingRoomsService.GetAllRoomAddressesFromActiveDirectory();

            foreach (var room in rooms)
            {
                try
                {
                    var appointments = MeetingRoomsService.GetAppointmentsForUser(room, 2);

                    if (appointments != null)
                    {
                        output += $"{room,-35}{appointments.Count,-45}{Environment.NewLine}";

                        if (appointments.Count > 0)
                        {
                            output += $"{Environment.NewLine}\tAppointments...{Environment.NewLine}";
                        }

                        foreach (var appointment in appointments)
                        {
                            // From the point of view of the meeting room, certian information is different than from the organizer's pov. Use this to get the organizer pov if so desired.
                            //var organizerAppointment =
                            //    MeetingRoomsService.GetAppointmentForUserByICalId(appointment.Organizer.Address, 2,
                            //        appointment.ICalUid);

                            _appointmentManager.AddOrUpdateAppointment(appointment);
                            output += $"\t-----------------------------------------------------------------{Environment.NewLine}";
                            output += $"\t           Subject: {appointment.Subject}{Environment.NewLine}";
                            output += $"\t        Start Time: {appointment.Start}{Environment.NewLine}";
                            output += $"\t          Location: {appointment.Location}{Environment.NewLine}"; 
                            output += $"\t          End Time: {appointment.End}{Environment.NewLine}";
                            output += $"\t           ICal Id: ...{appointment.ICalUid.Substring(appointment.ICalUid.Length - 30)}{Environment.NewLine}";
                            output += $"\t         Unique Id: ...{appointment.Id.UniqueId.Substring(appointment.Id.UniqueId.Length - 30)}{Environment.NewLine}";
                            output += $"\t         Change Id: {appointment.Id.ChangeKey}{Environment.NewLine}";
                            output += $"\t         Organizer: {appointment.Organizer}{Environment.NewLine}";
                            output += $"\tRequired Attendees: {GetAttendeeNamesAndResponse(appointment.RequiredAttendees)}{Environment.NewLine}";
                            output += $"\tOptional Attendees: {GetAttendeeNamesAndResponse(appointment.OptionalAttendees)}{Environment.NewLine}";
                            output += $"\t-----------------------------------------------------------------{Environment.NewLine}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex + Environment.NewLine);
                }
            }

            Console.WriteLine(output);
        }

        private static string GetAttendeeNamesAndResponse(AttendeeCollection attendees)
        {
            var nameList = new List<string>();
            foreach (var attendee in attendees)
            {
                var s = attendee.Name + " -> ";

                if (attendee.ResponseType.HasValue)
                {
                    s += attendee.ResponseType.Value;
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
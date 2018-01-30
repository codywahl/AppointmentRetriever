using AppointmentRetriever.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace AppointmentRetriever
{
    public class Program
    {
        private static AppointmentManager _appointmentManager;

        private static void Main(string[] args)
        {
            _appointmentManager = new AppointmentManager();

            GetMeetings();
            Console.WriteLine("Waiting...");
            Console.ReadKey();

            //CancelMeeting();

            GetMeetings();
            Console.WriteLine("Waiting...");
            Console.ReadKey();
        }

        private static void CancelMeeting()
        {
            MeetingRoomsService.CancelAppointment(@"AQMkAGMzZjAzMTI4LTIzODQtNDdhMi1hNTFlLWUzOGE0NTk2NTlmMABGAAADrvpyiRwtfEOF5Q/Zf/He0wcARDWNuRjujk23tS/kvy7RNwAAAw8AAABENY25GO6OTbe1L+S/LtE3AAADMAAAAA==", "Cuz i want to.");
        }

        private static void GetMeetings()
        {
            Console.WriteLine($"{"Room",-35}{"Appointment Count",-45}");
            Console.WriteLine($"-----------------------------------------------------------------------{Environment.NewLine}");

            var rooms = MeetingRoomsService.GetAllRoomAddressesFromActiveDirectory();

            foreach (var room in rooms)
            {
                try
                {
                    var appointments = MeetingRoomsService.GetAppointmentsForUser(room, 2);

                    if (appointments != null)
                    {
                        Console.WriteLine($"{room,-35}{appointments.Count,-45}");

                        if (appointments.Count > 0)
                        {
                            Console.WriteLine($"{Environment.NewLine}\tAppointments...");
                        }

                        foreach (var appointment in appointments)
                        {
                            // From the point of view of the meeting room, certian information is different than from the organizer's pov. Use this to get the organizer pov if so desired. 
                            //var organizerAppointment =
                            //    MeetingRoomsService.GetAppointmentForUserByICalId(appointment.Organizer.Address, 2,
                            //        appointment.ICalUid);

                            _appointmentManager.AddOrUpdateAppointment(appointment);
                            Console.WriteLine($"\t-----------------------------------------------------------------");
                            Console.WriteLine($"\t           Subject: {appointment.Subject}");
                            Console.WriteLine($"\t        Start Time: {appointment.Start}");
                            Console.WriteLine($"\t          End Time: {appointment.End}");
                            Console.WriteLine($"\t           ICal Id: ...{appointment.ICalUid.Substring(appointment.ICalUid.Length - 30)}");
                            Console.WriteLine($"\t         Unique Id: ...{appointment.Id.UniqueId.Substring(appointment.Id.UniqueId.Length - 30)}");
                            Console.WriteLine($"\t         Change Id: {appointment.Id.ChangeKey}");
                            Console.WriteLine($"\t         Organizer: {appointment.Organizer}");
                            Console.WriteLine($"\tRequired Attendees: {GetAttendeeNamesAndResponse(appointment.RequiredAttendees)}");
                            Console.WriteLine($"\tOptional Attendees: {GetAttendeeNamesAndResponse(appointment.OptionalAttendees)}");
                            Console.WriteLine($"\t-----------------------------------------------------------------");
                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex + Environment.NewLine);
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Appointment Manager Count: {_appointmentManager.GetAppointmentCount()}");
            Console.WriteLine();
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
            //var nameList = attendees.Select(attendee => attendee.Name + " -> " + attendee.ResponseType.Value ?? "test"  ).ToList();

            return string.Join(",", nameList);
        }
    }
}
using AppointmentRetriever.Services;
using System;

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
                            _appointmentManager.AddOrUpdateAppointment(appointment);
                            Console.WriteLine($"\t-----------------------------------------------------------------");
                            Console.WriteLine($"\t       Subject: {appointment.Subject}");
                            Console.WriteLine($"\t    Start Time: {appointment.Start}");
                            Console.WriteLine($"\t      End Time: {appointment.End}");
                            Console.WriteLine($"\t       ICal Id: {appointment.ICalUid}");
                            Console.WriteLine($"\t     Unique Id: {appointment.Id.UniqueId}");
                            Console.WriteLine($"\t     Change Id: {appointment.Id.ChangeKey}");
                            Console.WriteLine($"\t     Organizer: {appointment.Organizer}");
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
    }
}
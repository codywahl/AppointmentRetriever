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
            Console.WriteLine($"{Environment.NewLine}{"Room",-35}{"Appointment Count",-45}");
            Console.WriteLine($"-----------------------------------------------------------------------{Environment.NewLine}");

            var rooms = MeetingRoomsService.GetAllRoomAddressesFromActiveDirectory();
            var allAppointments = new List<Appointment>();

            foreach (var room in rooms)
            {
                try
                {
                    var roomAppointments = MeetingRoomsService.GetAppointmentsForUser(room, 2);
                    if (roomAppointments != null)
                    {
                        allAppointments.AddRange(roomAppointments);
                        Console.WriteLine($"{room,-35}{roomAppointments.Count,-45}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex + Environment.NewLine);
                }
            }

            if (allAppointments.Count > 0)
            {
                _appointmentManager.AddOrUpdateAppointments(allAppointments);
                Console.WriteLine(_appointmentManager.AppointmentsToString());
            }
        }

        #endregion Private Helpers
    }
}
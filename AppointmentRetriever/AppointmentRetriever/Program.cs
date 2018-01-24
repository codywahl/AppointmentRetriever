using System;

namespace AppointmentRetriever
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"{"Room",-35}{"Appointment Count",-45}");
            Console.WriteLine($"-----------------------------------------------------------------------{Environment.NewLine}");

            var rooms = ExchangeWebServicesWrapper.GetAllRoomAddressesFromActiveDirectory();

            foreach (var room in rooms)
            {
                try
                {
                    var appointments = ExchangeWebServicesWrapper.GetAppointmentsForUser(room, 2);

                    if (appointments != null)
                    {
                        Console.WriteLine($"{room,-35}{appointments.Count,-45}");

                        if (appointments.Count > 0)
                        {
                            Console.WriteLine($"{Environment.NewLine}\tAppointments...");
                        }

                        foreach (var appointment in appointments)
                        {
                            Console.WriteLine($"\t-----------------------------------------------------------------");
                            Console.WriteLine($"\tSubject:    {appointment.Subject}");
                            Console.WriteLine($"\tStart Time: {appointment.Start}");
                            Console.WriteLine($"\tEnd Time:   {appointment.End}");
                            Console.WriteLine($"\t-----------------------------------------------------------------");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex + Environment.NewLine);
                }
            }

            Console.ReadKey();
        }
    }
}
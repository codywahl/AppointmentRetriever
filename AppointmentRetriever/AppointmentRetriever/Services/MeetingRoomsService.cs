﻿using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.Linq;
using System.Net;

namespace AppointmentRetriever.Services
{
    public class MeetingRoomsService
    {
        #region Fields

        private static readonly string LdapAddress = ConfigurationManager.AppSettings["LDAPAddress"];
        private static readonly string ExchangeUserName = ConfigurationManager.AppSettings["ExchangeUserName"];
        private static readonly string ExchangeUserPassword = ConfigurationManager.AppSettings["ExchangeUserPassword"];
        private static readonly string ExchangeUserDomain = ConfigurationManager.AppSettings["ExchangeUserDomain"];
        private static readonly string ExchangeUrl = ConfigurationManager.AppSettings["ExchangeURL"];
        private static readonly string ExchangeVersion = ConfigurationManager.AppSettings["ExchangeVersion"];

        private const string LdapFilter = "(&(&(&(mailNickname=*)(objectcategory=person)(objectclass=user)(msExchRecipientDisplayType=7))))";

        #endregion Fields

        public static List<string> GetAllRoomAddressesFromActiveDirectory()
        {
            var directoryEntry = new DirectoryEntry(LdapAddress, ExchangeUserName, ExchangeUserPassword);
            var directorySearcher = new DirectorySearcher(directoryEntry)
            {
                Filter = LdapFilter
            };

            directorySearcher.PropertiesToLoad.Add("sn");
            directorySearcher.PropertiesToLoad.Add("mail");

            return (from SearchResult searchResult in directorySearcher.FindAll() select searchResult.Properties["mail"][0].ToString()).ToList();
        }

        public static List<Appointment> GetAppointmentsForUser(string mailboxToAccess, int numberOfMonths)
        {
            if (string.IsNullOrWhiteSpace(mailboxToAccess))
            {
                throw new ArgumentNullException(nameof(mailboxToAccess));
            }

            try
            {
                var service = GetExchangeService();
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxToAccess);

                var startDate = DateTime.Today;
                var endDate = DateTime.Today.AddMonths(numberOfMonths);
                var psPropSet = new PropertySet(BasePropertySet.FirstClassProperties);
                var calendarView = new CalendarView(startDate, endDate) { PropertySet = psPropSet };
                var calendarFolderId = new FolderId(WellKnownFolderName.Calendar, mailboxToAccess);

                var roomAppointments = service.FindAppointments(calendarFolderId, calendarView).ToList();

                if (roomAppointments.Count > 0)
                {
                    service.LoadPropertiesForItems(roomAppointments, psPropSet);
                }

                return roomAppointments;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine();
                return null;
            }
        }

        public static void CancelAppointment(string meetingId, string messageBody = "Cancelled by RightCrowd", bool isReadReceiptRequested = false)
        {
            var service = GetExchangeService();

            try
            {
                var itemId = new ItemId(meetingId);
                var appointment = Appointment.Bind(service, itemId);
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, appointment.Organizer.Address);

                appointment.Delete(DeleteMode.MoveToDeletedItems, SendCancellationsMode.SendOnlyToAll);
                //var cancelResults = appointment.CancelMeeting(messageBody);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        #region Private Helpers

        private static ExchangeService GetExchangeService(bool traceEnabled = false)
        {
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

            var version = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2013_SP1;

            if (!string.IsNullOrWhiteSpace(ExchangeVersion) && ExchangeVersion.Equals("2013"))
            {
                version = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2013;
            }

            var service = new ExchangeService(version)
            {
                TraceEnabled = traceEnabled,
                TraceEnablePrettyPrinting = traceEnabled,
                Credentials = new NetworkCredential(ExchangeUserName, ExchangeUserPassword, ExchangeUserDomain),
                Url = new Uri(ExchangeUrl)
            };

            if (traceEnabled)
            {
                service.TraceFlags = TraceFlags.All;
            }

            // Autodiscover does not work on our test exchange server
            // Best practice is to use the auto-discover, according to MS, but its not neccessary.
            // Calling autodiscover after setting the url in the constructor should, in theory, overwrite it, if it works. If not, oh well.
            //try
            //{
            //    service.AutodiscoverUrl($"{ExchangeUserName}@{ExchangeUserDomain}");
            //}
            //catch (AutodiscoverLocalException)
            //{
            //    //Autodiscover failed, use the previously set url.
            //}

            return service;
        }

        private static Appointment GetAppointmentForUserOnDateByICalId(string mailboxToAccess, DateTime startDate, DateTime endDate, string iCalId)
        {
            var service = GetExchangeService();

            try
            {
                var calendarView = new CalendarView(startDate, endDate);

                var calendarFolderId = new FolderId(WellKnownFolderName.Calendar, mailboxToAccess);
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxToAccess);

                var foundAppointments = service.FindAppointments(calendarFolderId, calendarView).ToList();

                if (foundAppointments.Count > 0)
                {
                    service.LoadPropertiesForItems(foundAppointments, PropertySet.FirstClassProperties);
                }

                return foundAppointments.FirstOrDefault(x => x.ICalUid.Equals(iCalId) && x.Start.Equals(startDate) && x.End.Equals(endDate));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine();
                return null;
            }
        }

        #endregion Private Helpers
    }
}
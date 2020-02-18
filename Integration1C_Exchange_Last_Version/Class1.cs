using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Security;
using System.Net;
using Newtonsoft.Json;

namespace Integration1C_Exchange_Last_Version
{
    [Guid("adae2508-e11b-4f82-867d-4fc0aa29906d")]
    internal interface IMyClass
    {
        [DispId(1)]
        string SetService(string userName, string password, string email, string domen);
        string GetRooms();
    }

    [Guid("b81ca92d-426a-441e-ac7b-ab5769f990a8"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMyEvents
    {
    }

    [Guid("1c128ea5-18cd-4da3-9777-700b59bfa9ea"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IMyEvents))]
    public class MyClass : IMyClass
    {

        private static ExchangeService serviceExchange { get; set; }

        public string SetService(string name, string passw, string email, string AD){

            //connect to Exchange Server
            string userName = name;
            string Email = email;
            string password = passw;
            string domen = AD;

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);

            SecureString secureString = new SecureString();
            for (int i = 0; i < password.Length; i++){
                secureString.AppendChar(password[i]);
            }

            service.Credentials = new NetworkCredential(userName, password, domen);
            service.AutodiscoverUrl(Email, RedirectionCallback); 
            
            serviceExchange = service;

            return "EWS Endpoint: "+service.Url;
        }

        public string SetServiceWithoutAutodiscover(string adminName, string nameFrom, string adminPassword, string AD)
        {

            //****************************ЗДЕСЬ МЕНЯТЬ ИМПЕРСОНИРОВАННУЮ УЧЕТКУ************************************
            //string adminName = "evoko";
            //string adminPassword = @"asdo;ai%^&UFSDAS*(2q1";
            //string AD = "pridex.local";
            //*******************************************************************************************************

            var ews = new ExchangeService(ExchangeVersion.Exchange2010_SP2, TimeZoneInfo.Utc);
         
            //защищаем пароль
            SecureString secureString = new SecureString();
            for (int i = 0; i < adminPassword.Length; i++){
                secureString.AppendChar(adminPassword[i]);
            }
            ews.Credentials = new NetworkCredential(adminName, adminPassword, AD);
            ews.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, nameFrom);
            ews.Url = new System.Uri("https://excas1.pridex.local/EWS/Exchange.asmx");
            ews.PreAuthenticate = true;

            serviceExchange = ews;

            return "EWS Endpoint: " + ews.Url;
        }

        static bool RedirectionCallback(string url){
            return url.ToLower().StartsWith("https://"); 
        }

        public string GetRooms(){

            string RoomsList = "";
            // Return all the room lists in the organization.
            // This method call results in a GetRoomLists call to EWS.
            EmailAddressCollection myRoomLists = serviceExchange.GetRoomLists();
            // Display the room lists.
            foreach (EmailAddress address in myRoomLists){
                EmailAddress myRoomList = address.Address;
                // This method call results in a GetRooms call to EWS.
                System.Collections.ObjectModel.Collection<EmailAddress> myRoomAddresses = serviceExchange.GetRooms(myRoomList);
                // Display the individual rooms.
                int countRoom = myRoomAddresses.Count;
                int count = 1;
                foreach (EmailAddress addressRoom in myRoomAddresses){
                    RoomsList = RoomsList + addressRoom.Address + (count < countRoom ? "," : "");
                    count += 1;
                }
            }
            return RoomsList;
        }

        public string GetSuggestedMeetingTimesAndFreeBusyInfo(string RoomAdress, string Start, string End, int meetingDuration){

            // Create a collection of attendees(only room) 
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            attendees.Add(new AttendeeInfo(){
                SmtpAddress = RoomAdress,
                AttendeeType = MeetingAttendeeType.Room
            });

            // Specify options to request free/busy information and suggested meeting times.
            AvailabilityOptions availabilityOptions = new AvailabilityOptions();
            availabilityOptions.GoodSuggestionThreshold = 49;
            availabilityOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
            availabilityOptions.MaximumSuggestionsPerDay = 5;
            availabilityOptions.MeetingDuration = meetingDuration;
            availabilityOptions.MinimumSuggestionQuality = SuggestionQuality.Good;
            availabilityOptions.DetailedSuggestionsWindow = new TimeWindow(Convert.ToDateTime(Start), Convert.ToDateTime(End));
            availabilityOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            //Return free-busy information and a set of suggested meeting times. 
            GetUserAvailabilityResults results = serviceExchange.GetUserAvailability(attendees,
                                                                             availabilityOptions.DetailedSuggestionsWindow,
                                                                             AvailabilityData.FreeBusyAndSuggestions,
                                                                             availabilityOptions);
            //Display suggested meeting times. 
            ReservationRoomInfo roomInfo = new ReservationRoomInfo();
            roomInfo.RoomName = attendees[0].SmtpAddress;

            foreach (Suggestion suggestion in results.Suggestions){

                foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions){
                    roomInfo.ListOfSuggestedMeetingTime.Add(timeSuggestion.MeetingTime.AddHours(3).ToShortTimeString()+","+
                                                            timeSuggestion.MeetingTime.AddHours(3).Add(TimeSpan.FromMinutes(availabilityOptions.MeetingDuration)).ToShortTimeString());
                }
            }

            int i = 0;
            //Display free-busy times.
            foreach (AttendeeAvailability availability in results.AttendeesAvailability){
                foreach (CalendarEvent calEvent in availability.CalendarEvents){
                    roomInfo.ListFreeBusyTimes.Add(calEvent.StartTime.AddHours(3).ToString()+","+
                                                            calEvent.EndTime.AddHours(3).ToString());
                }
                i++;
            }
            return JsonConvert.SerializeObject(roomInfo,Formatting.Indented);
        }

        public void CreateAppointment(string userName, string Start, string End, string roomAlias, string roomName) {

            try{
                //create new appointment
                Appointment appointment = new Appointment(serviceExchange);

                //set properties on the appointment
                appointment.Subject = "Быстрое бронирование из Телеграм";
                appointment.Body = "Встреча "+ userName;
                appointment.Start = Convert.ToDateTime(Start).AddHours(-3); 
                appointment.End = Convert.ToDateTime(End).AddHours(-3);
                appointment.Location = roomAlias; // "Think Tank";
                appointment.RequiredAttendees.Add(roomName); // "ThinkTank@pridex.local"

                //save the appointment
                appointment.Save(SendInvitationsMode.SendOnlyToAll);
                appointment.Update(ConflictResolutionMode.AutoResolve);

                Console.WriteLine("Appointment sucsess! Time start {0}, time end {1}", appointment.Start, appointment.End);

            }catch(Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        //class for serialization in json
        class ReservationRoomInfo {

            public string RoomName;
            public List<string> ListOfSuggestedMeetingTime = new List<string>();
            public List<string> ListFreeBusyTimes = new List<string>();
        }

    }
}

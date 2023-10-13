using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;


namespace Excel
{
    public class Oauth2ExampleAndGoogleCalender
    {
        public  async Task<UserCredential> GetCredentialAsync(string[] scopes)
        {
            UserCredential credential;

            using (var stream = new FileStream("C:\\Users\\ehb.it-minds.dk\\calender\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
            }

            return credential;
        }

        public  async Task ListEventsAsync(CalendarService service)
        {
            var request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.MaxResults = 10;
            request.SingleEvents = true;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            var events = await request.ExecuteAsync();

            if (events.Items != null && events.Items.Count > 0)
            {
                Console.WriteLine("Upcoming 10 events:");
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (string.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine($"{when} - {eventItem.Summary}");
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
        }

        public void CreateAnEvent(Event newEvent, CalendarService service)
        {
            EventsResource.InsertRequest request = service.Events.Insert(newEvent, "primary");
            Event createdEvent = request.Execute();
            Console.WriteLine($"Event created: {createdEvent.HtmlLink}");
        }
    }
}

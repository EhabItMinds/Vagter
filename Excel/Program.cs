using ClosedXML.Excel;
using Excel;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Net;


public class Program
{
    public static async Task Main(string[] args)
    {
        Oauth2ExampleAndGoogleCalender googleCalender = new Oauth2ExampleAndGoogleCalender();
        EventsToTheCalender toTheCalender = new EventsToTheCalender();
        IReading readingFromExel = new ReadingFromExcel();
        string excelFilePath = "C:\\Users\\ehb.it-minds.dk\\Desktop\\Vagtskema Svømmestadion.xlsx";
        string targetName = "";
        List<string> namesOfMedArbejder = new List<string>()
        {
            "ehab",
            "Signe",
        };
        List<Vagt> vagter;
        List<Event> events;

        while (!namesOfMedArbejder.Contains(targetName))
        {
            Console.WriteLine("Skrive medarbejderen navne for at finde hvornår personen skal arbejde :) ");
            targetName = Console.ReadLine();
            if (!namesOfMedArbejder.Contains(targetName))
                Console.WriteLine("The name is not a coworkername!");
        }
        vagter = readingFromExel.ReadData(targetName, excelFilePath);


        events = toTheCalender.events(vagter);

        string[] scopes = { CalendarService.Scope.Calendar };
        UserCredential credential = await googleCalender.GetCredentialAsync(scopes);

        if (credential != null)
        {
            CalendarService service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Vagter",
            });
            foreach (Event @event in events)
            {
                googleCalender.CreateAnEvent(@event, service);
            }



            await googleCalender.ListEventsAsync(service);


        }
        else
        {
            Console.WriteLine("No credentials found.");
        }


    }

}
using Google.Apis.Calendar.v3.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal class EventsToTheCalender
    {

        public List<Event> events(List<Vagt> vagter)
        {
            List<Event> events = new List<Event>();
            // Parse the date
            foreach (var vagt in vagter)
            {
                if (vagt.dato != "")
                {
                    DateTime date = Convert.ToDateTime(vagt.dato);

                    // Parse the start and end times
                    string[] timeParts = vagt.hours.Split('-');
                    string startTimeString = timeParts[0];
                    string endTimeString = timeParts[1];
                    if (startTimeString.Contains("."))
                    {
                        startTimeString = startTimeString.Replace(".", ":");
                    }

                    if (endTimeString.Contains("."))
                    {
                        endTimeString = endTimeString.Replace(".", ":");
                    }


                    // Parse the start and end times as TimeSpan
                    TimeSpan startTime = TimeSpan.Parse(startTimeString);
                    TimeSpan endTime = TimeSpan.Parse(endTimeString);

                    // Combine date and start time to create the start DateTime
                    DateTime startDateTime = date.Add(startTime);

                    // Combine date and end time to create the end DateTime
                    DateTime endDateTime = date.Add(endTime);

                    // Print the results
                    Console.WriteLine("Start Date and Time: " + startDateTime);
                    Console.WriteLine("End Date and Time: " + endDateTime);

                    events.Add(new Event
                    {
                        Summary = "Livredder ved Aarhus Svømmestadion",
                        Location = "F. Vestergaards Gade 5, 8000 Aarhus C",
                        Description = "Redde liv",
                        Start = new EventDateTime
                        {
                            DateTime = startDateTime,
                            TimeZone = "Europe/Copenhagen"
                        },
                        End = new EventDateTime
                        {
                            DateTime = endDateTime,
                            TimeZone = "Europe/Copenhagen"
                        },
                        Recurrence = /*new List<string> { "RRULE:FREQ=DAILY;COUNT=2" }*/ null,
                        Attendees = new List<EventAttendee>
                        {
                            //new EventAttendee { Email = "lpage@example.com" },
                            //new EventAttendee { Email = "sbrin@example.com" }
                        },
                        Reminders = new Event.RemindersData
                        {
                            UseDefault = false,
                            Overrides = new List<EventReminder>
                    {
                        new EventReminder { Method = "email", Minutes = 24 * 60 },
                        new EventReminder { Method = "popup", Minutes = 10 }
                    }
                        }
                    });
                }

            }
            return events;
        }
    }
}

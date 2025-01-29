using Microsoft.Office.Interop.Outlook;
using System;

using Microsoft.Office.Interop.Outlook;
using System;

class Program
{
    static void Main(string[] args)
    {
        Application outlookApp = new Application();
        NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

        // Get default calendar
        MAPIFolder calendarFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        Items calendarItems = calendarFolder.Items;

        // Sort and filter calendar appointments for 2025
        //calendarItems.Sort("[Start]");
        //string filter = "[Start] >= '01/01/2025'";
        Items filteredCalendarItems = calendarItems;//.Restrict(filter);

        Console.WriteLine("Accepted or Tentative Calendar Appointments:");
        foreach (AppointmentItem appointment in filteredCalendarItems)
        {
            Console.WriteLine($"- {appointment.Subject} ({appointment.Start})");
            if (appointment.IsRecurring)
                Console.WriteLine("  [RECURRING]");
            if (appointment.MeetingStatus == OlMeetingStatus.olMeetingCanceled || appointment.MeetingStatus == OlMeetingStatus.olMeetingReceivedAndCanceled)
                Console.WriteLine("  Ol' Meeting's cancelled y'all");
        }

        // Get Inbox to find unaccepted meeting requests
        MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        Items inboxItems = inboxFolder.Items;
        inboxItems = inboxItems.Restrict("[MessageClass] = 'IPM.Schedule.Meeting.Request'");

        Console.WriteLine("\nPending Meeting Invitations:");
        foreach (object item in inboxItems)
        {
            if (item is MeetingItem meetingRequest)
            {
                try
                {
                    AppointmentItem tentativeMeeting = meetingRequest.GetAssociatedAppointment(false);
                    if (tentativeMeeting != null)
                    {
                        Console.WriteLine($"- {tentativeMeeting.Subject} (Tentative - {tentativeMeeting.Start})");
                        if (tentativeMeeting.IsRecurring)
                            Console.WriteLine("  [RECURRING]");
                        if (tentativeMeeting.MeetingStatus == OlMeetingStatus.olMeetingCanceled || tentativeMeeting.MeetingStatus == OlMeetingStatus.olMeetingReceivedAndCanceled)
                            Console.WriteLine("  Ol' Meeting's cancelled y'all");
                    }
                }
                catch
                {
                    Console.WriteLine($"- {meetingRequest.Subject} (Pending - No date found)");
                }
            }
        }
    }
}

/*
 * 
 * 
Subject: Nieuwe location vanuit sync api onder chargepointupdate zonder error opslaan
Start: 26-2-2021 10:00:00
End: 26-2-2021 11:30:00

Subject: Code review
Start: 26-2-2021 11:30:00
End: 26-2-2021 13:30:00

Subject: 28248 User invite flow tracking issue
Start: 26-2-2021 14:00:00
End: 26-2-2021 17:00:00

Subject: code review
Start: 3-3-2021 10:30:00
End: 3-3-2021 11:00:00
*/
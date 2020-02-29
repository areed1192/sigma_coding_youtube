Option Explicit

Sub WorkingWithAppointmentItems()

'Declare our Variables
Dim oLookApptItem As AppointmentItem
Dim oLookRecipients As Recipients
Dim oLookRecipient As Recipient
Dim oLookMeetingItem As MeetingItem
Dim oLookRecPattern As RecurrencePattern

'Let's work an email
Set oLookApptItem = Application.CreateItem(olAppointmentItem)

'Define some attributes about our new meeting, the first beeing the meeting status
oLookApptItem.MeetingStatus = olMeeting

    'olMeeting                       1   The meeting has been scheduled.
    'olMeetingCanceled               5   The scheduled meeting has been cancelled.
    'olMeetingReceived               3   The meeting request has been received.
    'olMeetingReceivedAndCanceled    7   The scheduled meeting has been cancelled but still appears on the user's calendar.
    'olNonMeeting                    0   An Appointment item without attendees has been scheduled. This status can be used to
    '                                    set up holidays on a calendar.

'Give it a subject line
oLookApptItem.Subject = "New Employee Meeting"

'Give it a location
oLookApptItem.Location = "Conference Room B"

'Give it a Start Time
oLookApptItem.Start = #1/28/2020 8:00:00 AM#

'Set the duration
oLookApptItem.Duration = 90

'Set the reminder time.
oLookApptItem.ReminderMinutesBeforeStart = 30

'Let's add some Recipients, this one will be Required.
Set oLookRecipient = oLookApptItem.Recipients.Add("Bob Gates")
    
    'Set the Receipient Type
    oLookRecipient.Type = olRequired

'Let's add some Recipients, this one will be Optional.
Set oLookRecipient = oLookApptItem.Recipients.Add("Bob Gates 2")

    'Set the Receipient Type
    oLookRecipient.Type = olOptional

'Let's add some Recipients, this one will be Our Resource.
Set oLookRecipient = oLookApptItem.Recipients.Add("Conference Room B")

    'Set the Receipient Type
    oLookRecipient.Type = olResource

'Let's add some Recipients, this one will be the Organizer.
Set oLookRecipient = oLookApptItem.Recipients.Add("Alex Reed")

    'Set the Receipient Type
    oLookRecipient.Type = olOrganizer
 
'Grab the Recurrence Pattern
Set oLookRecPattern = oLookApptItem.GetRecurrencePattern

    'Let's have it reoccur monthly.
    oLookRecPattern.RecurrenceType = olRecursMonthly
    
    'Define the start date
    oLookRecPattern.PatternStartDate = #1/26/2020#
    
    'Define the end date
    oLookRecPattern.PatternEndDate = #1/26/2021#

'Set the Body
oLookApptItem.Body = "Make sure to attend this meeting!"
oLookApptItem.BodyFormat = olFormatHTML

'Save the Appointment Item.
oLookApptItem.Save

'Display it.
oLookApptItem.Display

End Sub
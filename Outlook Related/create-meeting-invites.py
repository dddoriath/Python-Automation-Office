import win32com.client as client
outlook = client.Dispatch("Outlook.Application")

def createMeeting():
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.Start = "2022-03-15 10:10" # yyyy-MM-dd hh:mm
    appt.Subject = "Subject of the meeting"
    appt.Duration = 60 # In minutes (60 Minutes)
    appt.Location = "Location Name"
    appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added

    appt.Recipients.Add("bobbie.wxy@gmail.com") # Don't end ; as delimiter
    appt.Save()
    #appt.Send() #send

appt = createMeeting()
namespace = outlook.GetNameSpace('MAPI')
drafts = namespace.GetDefaultFolder(16)
messages = list(drafts.Items)

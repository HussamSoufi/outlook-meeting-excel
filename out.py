import win32com.client
import xlrd
from datetime import datetime, date, time, timedelta


workbook = xlrd.open_workbook('myfile.xlsx')
worksheet = workbook.sheet_by_name('Sheet2')
num_rows = worksheet.nrows - 1
num_cols = worksheet.ncols


def addevent(start, subject, duration):
    oOutlook = win32com.client.Dispatch("Outlook.Application")
    appointment = oOutlook.CreateItem(1)  # 1=outlook appointment item
    appointment.Start = start
    appointment.Subject = subject
    appointment.Duration = duration
    # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
    appointment.MeetingStatus = 1
    #appointment.Location = 'Berlin'
    appointment.Recipients.Add("whatever@gmail.com")
    appointment.Recipients.ResolveAll()
    appointment.ReminderSet = True
    appointment.ReminderMinutesBeforeStart = 1
    appointment.Save()
    appointment.Send()
    return


for column in range(1, num_cols):
    for row in range(6, 17):
        try:
            exDatum = worksheet.cell_value(row, column)
            y, m, d, h, i, s = xlrd.xldate_as_tuple(
                float(exDatum), workbook.datemode)
            Datum = "{0}-{1}-{2}".format(y, m, d)
            start = Datum
            subject = worksheet.cell_value(19, column)
            if worksheet.cell_value(28, column) == 7.50:
                duration = 1440
                addevent(start, subject, duration)
            if worksheet.cell_value(28, column) == 15.00:
                duration = 1440*2
                addevent(start, subject, duration)
            elif worksheet.cell_value(28, column) == 22.50:
                duration = 1440*3
                addevent(start, subject, duration)
        except ValueError:
            pass

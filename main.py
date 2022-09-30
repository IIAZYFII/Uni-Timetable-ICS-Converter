from ics import Calendar, Event
import pandas
import openpyxl
import datetime
from datetime import timedelta

path = "Path\\to\\excelfile"
c = Calendar()
number_rows = len(pandas.read_excel(path))
data = pandas.read_excel(path)


def get_day_index(day):
    match day:
        case 'Monday':
            return 0
        case 'Tuesday':
            return 1
        case 'Wednesday':
            return 2
        case 'Thursday':
            return 3
        case 'Friday':
            return 4


for i in range(number_rows):
    lecture_name = data['Activity Name / Enw’r Gweithgaredd'][i]
    day = data['Day / Dydd'][i]
    start_time = data['Start time / Amser dechrau'][i]
    finish_time = data['End Time'][i]
    date_range = data['Date range'][i]
    location = data['Locations / Lleoliad'][i]

    starting_date = date_range[:-13]
    begin_day = int(starting_date[:-8])
    begin_month = int(starting_date[3:5])
    begin_year = int(starting_date[6:])

    begin_date = datetime.date(begin_year, begin_month, begin_day)

    finishing_date = date_range[13:]
    end_day = int(finishing_date[:-8])
    end_month = int(finishing_date[3:5])
    end_year = int(finishing_date[6:])

    end_date = datetime.date(end_year, end_month, end_day)

    timezone_change = datetime.date(2022, 10, 30)

    weeks = int(((end_date - begin_date).days) / 7)

    for i in range(weeks):
        week_commencing = begin_date + timedelta(days=i * 7)
        final_date = week_commencing + timedelta(days=(get_day_index(day)))

        if final_date < timezone_change:
            begin_time = datetime.time((int(start_time[0:2])) - 1)
            end_time = datetime.time((int(finish_time[0:2])) - 1)
        else:
            begin_time = datetime.time(int(start_time[0:2]))
            end_time = datetime.time(int(finish_time[0:2]))

        begin = (final_date.strftime("%Y/%m/%d") + ' ' + begin_time.strftime("%H:%M:%S"))
        end = (final_date.strftime("%Y/%m/%d") + ' ' + end_time.strftime("%H:%M:%S"))

        e = Event()
        e.name = lecture_name
        e.location = location
        e.begin = begin
        e.end = end
        c.events.add(e)

with open('myTimetable.ics', 'w') as f:
    f.writelines(c.serialize_iter())

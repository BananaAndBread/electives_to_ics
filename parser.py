import re
from datetime import datetime
from ics import Calendar, Event
import pyexcel
import requests


def download_spreadsheet():
    def getFilename_fromCd(cd):
        """
        Get filename from content-disposition
        """

    file_id = "1h0VhA48io0Z345gPtXVr7S1OTKolrA3JrjMeFhFLsQI"
    url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
    r = requests.get(url, allow_redirects=True)
    filename = getFilename_fromCd(r.headers.get('content-disposition'))
    open("Electives Schedule Spring 2020 Bachelors.xlsx", 'wb').write(r.content)

def find_electives_in_row(electives_columns, row):
    electives_in_row = list()
    for column in electives_columns:

        if row[column]!="":
            electives_in_row.append({"name": electives_columns[column], "description":row[column]})
    if len(electives_in_row)!=0:
        return electives_in_row
    else:
        return None

def get_electives_columns(electives, first_row):
    columns = {}
    for elective_index in range(len(first_row)):
        if first_row[elective_index] in electives:
            columns[elective_index] = first_row[elective_index]
    return columns

download_spreadsheet()
sheet = pyexcel.get_sheet(file_name="Electives Schedule Spring 2020 Bachelors.xlsx").to_array()
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
            'Sunday']
c = Calendar()
first_row = sheet[0]

print("Choose electives from the list")
print("(As an example, if you want to choose"
      " 'Advanced agile software design' and 'Economics of Entrepreneurship' "
      "type '3,6') ")

for i in range(3, len(first_row)):
    print(f"{i-2}) {first_row[i]}")
print("Waiting for the input ...")
chosen_electives = input()
chosen_electives = chosen_electives.split(",")
electives = [first_row[int(i)+2] for i in chosen_electives]
electives_columns = get_electives_columns(electives, first_row)

for i in sheet:

    if i[0]!="Date":
        if i[0]!="":
            if isinstance(i[0], datetime):
                date_object= i[0]
            else:
                temp = str(i[0]).split(" ")[0]
                date_object = datetime.strptime(temp, '%d/%m/%Y')

        electives_found = find_electives_in_row(electives_columns, i)
        if electives_found:
            time = i[2]
            begin_hour = int(time.split("-")[0].split(":")[0])
            begin_minute = int(time.split("-")[0].split(":")[1])
            end_hour = int(time.split("-")[1].split(":")[0])
            end_minute = int(time.split("-")[1].split(":")[1])

            begin = date_object.replace(hour=begin_hour, minute=begin_minute)
            end = date_object.replace(hour=end_hour, minute=end_minute)
            for elective in electives_found:
                e = Event()
                e.name = elective["name"]
                e.description = elective["description"]
                e.begin = begin
                e.end = end
                c.events.add(e)


print("You have chosen the following electives:")
for elective in electives:
    print(f"* {elective}")
print(f"You got a calendar 'Electives.ics' with {len(c.events)} events.")


open('Electives.ics', 'w').writelines(c)





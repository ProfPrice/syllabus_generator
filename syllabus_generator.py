import json
import csv
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import filedialog
import os
from dateutil.easter import easter

def nth_weekday(n, weekday, month, year):
    """
    Returns the date of the nth occurrence of the weekday in the specified month and year.
    n: nth occurrence (e.g., 1 for first, 2 for second, etc.)
    weekday: 0 for Monday, 1 for Tuesday, ..., 6 for Sunday
    month: 1 for January, 2 for February, ..., 12 for December
    year: the year
    """
    count = 0
    day = 1
    while count < n:
        if date(year, month, day).weekday() == weekday:
            count += 1
        if count < n:
            day += 1
    return date(year, month, day)

""" # Open file dialog to select the JSON file
root = tk.Tk()
root.withdraw()  # Hide the main window
default_dir = os.path.join(os.getcwd(), 'JSON')
json_file_path = filedialog.askopenfilename(initialdir=default_dir, title="Select the JSON course data file", filetypes=[("JSON files", "*.json")]) """
# Hardcoded relative file path
json_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'JSON', 'mme2259.json')


# Derive the output filename from the input JSON filename
output_filename = os.path.basename(json_file_path).replace('.json', '.csv')
output_path = os.path.join('Output', output_filename)

# Load course data from the selected JSON file
with open(json_file_path, 'r') as file:
    data = json.load(file)

# Define activity types
lesson_entry_type = "Class section - Lecture"
lab_entry_type = "Class section - Lab"
tutorial_entry_type = "Tutorial"

# Extract term dates and other important dates
term_dates = {key: datetime.strptime(value, '%Y-%m-%d').date() for key, value in data['TermDates'].items()}
unavailable_dates = [datetime.strptime(date, '%Y-%m-%d').date() for date in data['UnavailableDates']]

# Extract the delivery option
delivery_option = data.get('DeliveryOption', 'Conventional')

# Calculate recurring holidays for the term year
year = term_dates['term_start_date'].year

# Canadian Thanksgiving - Second Monday of October
thanksgiving = nth_weekday(2, 0, 10, year)  # 0 for Monday
if thanksgiving < term_dates['term_start_date']:
    thanksgiving = nth_weekday(2, 0, 10, year + 1)

# Canada's Day of Truth and Reconciliation - September 30th
truth_and_reconciliation = date(year, 9, 30)
if truth_and_reconciliation.weekday() == 5:  # Saturday
    truth_and_reconciliation -= timedelta(days=1)
elif truth_and_reconciliation.weekday() == 6:  # Sunday
    truth_and_reconciliation += timedelta(days=1)
if truth_and_reconciliation < term_dates['term_start_date']:
    truth_and_reconciliation = datetime.date(year + 1, 9, 30)
    if truth_and_reconciliation.weekday() == 5:  # Saturday
        truth_and_reconciliation -= timedelta(days=1)
    elif truth_and_reconciliation.weekday() == 6:  # Sunday
        truth_and_reconciliation += timedelta(days=1)

# Good Friday - Friday before Easter Sunday
easter_sunday = easter(year)
good_friday = easter_sunday - timedelta(days=2)
if good_friday < term_dates['term_start_date']:
    easter_sunday = easter(year + 1)
    good_friday = easter_sunday - timedelta(days=2)

# Family Day - Third Monday of February
family_day = nth_weekday(3, 0, 2, year)  # 0 for Monday
if family_day < term_dates['term_start_date']:
    family_day = nth_weekday(3, 0, 2, year + 1)

# Add the holidays to the unavailable dates
unavailable_dates.extend([thanksgiving, truth_and_reconciliation, good_friday, family_day])

# Ensure term start date, term end date, reading week start, and unavailable dates are not in the past
current_date = datetime.now().date()
one_month_ago = current_date - timedelta(days=30)  # Approximate a month as 30 days

if term_dates['term_start_date'] < one_month_ago:
    raise ValueError("The term start date is in the past. Please provide a valid future date.")

if term_dates['term_end_date'] < current_date:
    raise ValueError("The term end date is in the past. Please provide a valid future date.")

if term_dates['reading_week_start'] < current_date:
    raise ValueError("The reading week start date is in the past. Please provide a valid future date.")

for date in unavailable_dates:
    if date < current_date:
        raise ValueError(f"The unavailable date {date} is in the past. Please provide a valid future date.")
    
# Print a summary of key dates and holidays
print("Summary of Key Dates and Holidays:")
print("----------------------------------")
print(f"Term Start Date: {term_dates['term_start_date']}")
print(f"Term End Date: {term_dates['term_end_date']}")
print(f"Reading Week Start: {term_dates['reading_week_start']}")
print(f"Unavailable Dates from JSON: {', '.join(map(str, data['UnavailableDates']))}")
print(f"Day of Truth and Reconciliation: {truth_and_reconciliation}")
print(f"Thanksgiving: {thanksgiving}")
print(f"Good Friday: {good_friday}")
print(f"Family Day: {family_day}")
print("----------------------------------\n")


# Extract topics
lecture_topics = data['LectureTopics']
tutorial_topics = data['TutorialTopics']
lab_topics = data['LabTopics']

# Initialize indices for topics
lecture_topic_index = 0
tutorial_topic_index = 0
lab_topic_index = 0

# Initialize storage for scheduled events
scheduled_activities = []

# Helper function to get the next available date for a given day of the week
def get_next_available_date(current_date, day_of_week, unavailable_dates):
    while True:
        if current_date.strftime('%A').upper() == day_of_week and current_date not in unavailable_dates:
            return current_date
        current_date += timedelta(days=1)

current_date = term_dates['term_start_date']
is_lecture_week = True  # Start with a lecture for the alternating option

while current_date <= term_dates['term_end_date']:
    if term_dates['reading_week_start'] <= current_date < term_dates['reading_week_start'] + timedelta(days=5) or current_date in unavailable_dates:
        current_date += timedelta(days=1)
        continue

   # Schedule lectures
    if delivery_option == 'Conventional' or (delivery_option == 'Alternating' and is_lecture_week):
        for lecture in data['Class section - Lecture']:
            if current_date.strftime('%A').upper() == lecture['day_of_week'] and lecture_topic_index < len(lecture_topics):
                start_time = lecture['start_time']
                duration = lecture['duration']
                location = lecture['location']
                topic = lecture_topics[lecture_topic_index]['Topic']
                end_time = (datetime.strptime(start_time, '%I:%M %p') + timedelta(hours=duration)).strftime('%I:%M %p')
                scheduled_activities.append((current_date, start_time, end_time, location, lesson_entry_type, topic))
                lecture_topic_index += 1

    def schedule_activities(data, delivery_option, is_lecture_week, topics, entry_type):
    # Check if the entry_type exists in the data dictionary
        if entry_type not in data:
            print(f"Warning: {entry_type} not found in the data. Skipping scheduling for this entry type.")
            return

        for topic in topics:
            for activity in data[entry_type]:
                section = activity['section']
                start_time = activity['start_time']
                duration = activity['duration']
                location = activity['location']
                end_time = (datetime.strptime(start_time, '%I:%M %p') + timedelta(hours=duration)).strftime('%I:%M %p')
                scheduled_activities.append((current_date, start_time, end_time, location, entry_type, topic + f" (Section {section})"))

    # Schedule labs
    if data['HasLabs'] and (delivery_option == 'Conventional' or (delivery_option == 'Alternating' and not is_lecture_week)):
        schedule_activities(data, delivery_option, is_lecture_week, lab_topics, lab_entry_type)

    # Schedule tutorials
    if data['HasTutorials']:
        schedule_activities(data, delivery_option, is_lecture_week, tutorial_topics, tutorial_entry_type)

    current_date += timedelta(days=1)
    if delivery_option == 'Alternating':
        is_lecture_week = not is_lecture_week # When the last lab has been scheduled keep this true in pertetuity.

# Sort the scheduled activities by date and start time
scheduled_activities.sort(key=lambda x: (x[0], datetime.strptime(x[1], '%I:%M %p')))

# Write the sorted activities to the CSV
with open(output_path, 'w', newline='') as csvfile:
    output_csv = csv.writer(csvfile)
    output_csv.writerow(["Date", "Start Time", "End Time", "Location", "Type", "Description"])
    for activity in scheduled_activities:
        output_csv.writerow(activity)


# Display summary
course_code = data['CourseCode']
deliverables = data['Deliverables']
total_weight = sum([deliverable['Weight'] for deliverable in deliverables])

print("\nCourse {course_code} Deliverables:")
print("+----------------------------------+--------+")
print("| Deliverable                      | Weight |")
print("+----------------------------------+--------+")
for deliverable in deliverables:
    name = deliverable['Name']
    weight = deliverable['Weight']
    print(f"| {name:30} | {weight:6}% |")
print("+----------------------------------+--------+")

# Summary and Warnings
print("\nSummary:")

# Calculate total number of lecture times available in the term
total_lecture_times = 0
current_date = term_dates['term_start_date']
while current_date <= term_dates['term_end_date']:
    for lecture in data['Lectures']:
        if current_date.strftime('%A').upper() == lecture['day_of_week']:
            total_lecture_times += 1
    current_date += timedelta(days=1)
print(f"Total lectures scheduled: {lecture_topic_index} topics/{total_lecture_times} lessons")

# Check for unused lecture times
if total_lecture_times > len(lecture_topics):
    print(f"Warning: There are {total_lecture_times - len(lecture_topics)} lecture times remaining in the term that have not been utilized with topics.")
elif total_lecture_times == len(lecture_topics):
    print("All lecture times in the term have been successfully utilized with topics.")
else:
    print(f"Warning: More topics provided ({len(lecture_topics)}) than available lecture times ({total_lecture_times}).")

print(f"Total tutorials scheduled (per section): {tutorial_topic_index}/{len(tutorial_topics)}")
print(f"Total labs scheduled (per section): {lab_topic_index}/{len(lab_topics)}")
print(f"Total weight of deliverables: {total_weight}%\n")

# Check for discrepancies in scheduled lectures vs. available topics
if lecture_topic_index < len(lecture_topics):
    print(f"Warning: Not all lecture topics were scheduled. {len(lecture_topics) - lecture_topic_index} topics remain unscheduled.")
elif lecture_topic_index > len(lecture_topics):
    print(f"Warning: More lectures were scheduled than available topics. {lecture_topic_index - len(lecture_topics)} extra lectures were scheduled without topics.")

# Check if total weight of deliverables is not 100%
if total_weight != 100:
    print(f"Warning: The total weight of deliverables is {total_weight}%, which is not equal to the expected 100%.")


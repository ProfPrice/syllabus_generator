import json
import csv
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
import os
from dateutil.easter import easter

# Open file dialog to select the JSON file
root = tk.Tk()
root.withdraw()  # Hide the main window
default_dir = os.path.join(os.getcwd(), 'JSON')
json_file_path = filedialog.askopenfilename(initialdir=default_dir, title="Select the JSON course data file", filetypes=[("JSON files", "*.json")])

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

# Calculate recurring holidays
year = term_dates['term_start_date'].year

# Canadian Thanksgiving (second Monday in October)
thanksgiving = datetime(year, 10, 1).date()
while thanksgiving.weekday() != 0:  # 0 represents Monday
    thanksgiving += timedelta(days=1)
thanksgiving += timedelta(weeks=1)

# Canada's Day of Truth and Reconciliation (September 30th, but observed on the preceding Friday if it falls on a weekend)
truth_and_reconciliation = datetime(year, 9, 30).date()
if truth_and_reconciliation.weekday() == 5:  # Saturday
    truth_and_reconciliation -= timedelta(days=1)
elif truth_and_reconciliation.weekday() == 6:  # Sunday
    truth_and_reconciliation -= timedelta(days=2)

# Good Friday (two days before Easter Sunday)
good_friday = easter(year) - timedelta(days=2)

# Family Day (third Monday in February for most provinces)
family_day = datetime(year, 2, 1).date()
while family_day.weekday() != 0:
    family_day += timedelta(days=1)
family_day += timedelta(weeks=2)

# Add these holidays to the unavailable dates list
unavailable_dates.extend([thanksgiving, truth_and_reconciliation, good_friday, family_day])

# Extract topics
lecture_topics = data['LectureTopics']
tutorial_topics = data['TutorialTopics']
lab_topics = data['LabTopics']

# Initialize indices for topics
lecture_topic_index = 0
tutorial_topic_index = 0
lab_topic_index = 0

# Helper function to get the next available date for a given day of the week
def get_next_available_date(current_date, day_of_week, unavailable_dates):
    while True:
        if current_date.strftime('%A').upper() == day_of_week and current_date not in unavailable_dates:
            return current_date
        current_date += timedelta(days=1)

# Open output CSV file for writing
with open('Output/scheduled_activities.csv', 'w', newline='') as csvfile:
    output_csv = csv.writer(csvfile)
    output_csv.writerow(['Date', 'Start Time', 'End Time', 'Location', 'Activity Type', 'Topic'])

    current_date = term_dates['term_start_date']
    while current_date <= term_dates['term_end_date']:
        if term_dates['reading_week_start'] <= current_date < term_dates['reading_week_start'] + timedelta(days=5) or current_date in unavailable_dates:
            current_date += timedelta(days=1)
            continue

        # Schedule lectures
        for lecture in data['Lectures']:
            if current_date.strftime('%A').upper() == lecture['day_of_week'] and lecture_topic_index < len(lecture_topics):
                start_time = lecture['start_time']
                duration = lecture['duration']
                location = lecture['location']
                topic = lecture_topics[lecture_topic_index]['Topic']
                output_csv.writerow([current_date, start_time, (datetime.strptime(start_time, '%H:%M') + timedelta(hours=duration)).time(), location, lesson_entry_type, topic])
                lecture_topic_index += 1

        # Schedule labs
        if data['HasLabs']:
            while lab_topic_index < len(lab_topics):
                sections_scheduled_for_labs = 0
                for lab in data['Labs']:
                    lab_date = get_next_available_date(term_dates['lab_start_date'], lab['day_of_week'], unavailable_dates)
                    if lab_date <= term_dates['term_end_date']:
                        start_time = lab['start_time']
                        duration = lab['duration']
                        location = lab['location']
                        section = lab['section']
                        topic = lab_topics[lab_topic_index]['Topic']
                        output_csv.writerow([lab_date, start_time, (datetime.strptime(start_time, '%H:%M') + timedelta(hours=duration)).time(), location, lab_entry_type, topic + " (Section " + section + ")"])
                        sections_scheduled_for_labs += 1
                if sections_scheduled_for_labs == len(data['Labs']):
                    lab_topic_index += 1

        # Schedule tutorials
        if data['HasTutorials']:
            while tutorial_topic_index < len(tutorial_topics):
                sections_scheduled = 0
                for tutorial in data['Tutorials']:
                    tutorial_date = get_next_available_date(term_dates['tutorial_start_date'], tutorial['day_of_week'], unavailable_dates)
                    if tutorial_date <= term_dates['term_end_date']:
                        start_time = tutorial['start_time']
                        duration = tutorial['duration']
                        location = tutorial['location']
                        section = tutorial['section']
                        topic = tutorial_topics[tutorial_topic_index]['Topic']
                        output_csv.writerow([tutorial_date, start_time, (datetime.strptime(start_time, '%H:%M') + timedelta(hours=duration)).time(), location, tutorial_entry_type, topic + " (Section " + section + ")"])
                        sections_scheduled += 1
                if sections_scheduled == len(data['Tutorials']):
                    tutorial_topic_index += 1

        current_date += timedelta(days=1)

# Display summary
course_code = data['CourseCode']
deliverables = data['Deliverables']
total_weight = sum([deliverable['Weight'] for deliverable in deliverables])

print(f"\nCourse {course_code} Deliverables:")
print("+----------------------------------+--------+")
print("| Deliverable                      | Weight |")
print("+----------------------------------+--------+")
for deliverable in deliverables:
    name = deliverable['Name']
    weight = deliverable['Weight']
    print(f"| {name:30} | {weight:6}% |")
print("+----------------------------------+--------+")

if total_weight == 100:
    print(f"\nThe total weight of all deliverables is 100%.")
else:
    print(f"\nWarning! The total weight of all deliverables is {total_weight}%, which deviates from the expected 100%.")

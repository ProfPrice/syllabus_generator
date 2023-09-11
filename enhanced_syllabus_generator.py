import json
import csv
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import filedialog
import os
from dateutil.easter import easter
from prettytable import PrettyTable

def print_warning(message):
    print(f"❌ Warning: {message}")


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
unavailable_dates = sorted([datetime.strptime(date, '%Y-%m-%d').date() for date in data['UnavailableDates']])


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
    
# Create a new table named dateSummaryTable
dateSummaryTable = PrettyTable()
dateSummaryTable.field_names = ["Description", "Date(s)"]
# Set alignment for columns
dateSummaryTable.align["Description"] = "l"
dateSummaryTable.align["Date(s)"] = "c"

# Add rows to the dateSummaryTable
dateSummaryTable.add_row(["Term Start Date", term_dates['term_start_date']])
dateSummaryTable.add_row(["Term End Date", term_dates['term_end_date']])
dateSummaryTable.add_row(["Reading Week Start", term_dates['reading_week_start']])
dateSummaryTable.add_row(["Unavailable Dates from JSON", ', '.join(map(str, data['UnavailableDates']))])
dateSummaryTable.add_row(["Day of Truth and Reconciliation", truth_and_reconciliation])
dateSummaryTable.add_row(["Thanksgiving", thanksgiving])
dateSummaryTable.add_row(["Good Friday", good_friday])
dateSummaryTable.add_row(["Family Day", family_day])

# Print the dateSummaryTable
print("Summary of Key Dates and Holidays:")
print(dateSummaryTable)


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
    # Check if current_date is within the reading week range
    is_within_reading_week = term_dates['reading_week_start'] <= current_date < term_dates['reading_week_start'] + timedelta(days=5)

    # Check if current_date is in the unavailable dates
    is_unavailable_date = current_date in unavailable_dates

    # Skip date if not available for scheduling
    if is_within_reading_week or is_unavailable_date:
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
            print_warning(f"{entry_type} not found in the data. Skipping scheduling for this entry type.")
            return

        for topic in topics:
            for activity in data[entry_type]:
                section = activity['section']
                start_time = activity['start_time']
                duration = activity['duration']
                location = activity['location']
                end_time = (datetime.strptime(start_time, '%I:%M %p') + timedelta(hours=duration)).strftime('%I:%M %p')
                # Construct the full topic description with the section
                topic_content = topic["Topic"]
                reference_content = topic.get("Reference", "")
                # Construct the full topic description with the section
                if reference_content:
                    full_topic_description = f"{topic_content} ({reference_content}) (Section {section})"
                else:
                    full_topic_description = f"{topic_content} (Section {section})"
                scheduled_activities.append((current_date, start_time, end_time, location, entry_type, full_topic_description))

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

print(f"\nCourse {course_code} Deliverables:")

deliverableTable = PrettyTable()
deliverableTable.field_names = ["Deliverable", "Weight"]
deliverableTable.align["Deliverable"] = "l"

for index, deliverable in enumerate(deliverables):
    name = deliverable['Name']
    weight = f"{deliverable['Weight']}%"
    
    # Check if it's the last iteration
    if index == len(deliverables) - 1:
        deliverableTable.add_row([name, weight], divider=True)
    else:
        deliverableTable.add_row([name, weight])

# Add the total weight to the table
deliverableTable.add_row(["TOTAL", f"{total_weight}%"])

print(deliverableTable)

# Check if total weight of deliverables is not 100%
if total_weight != 100:
    print_warning(f"The total sum of deliverables is {total_weight}%, which is not equal to the expected 100%.")
else:
    print(f"✅ The total sum of deliverables is {total_weight}%")

# Calculate total number of lecture times available in the term
total_lecture_times = 0
current_date = term_dates['term_start_date']

while current_date <= term_dates['term_end_date']:
    
    # Check if current_date is within the reading week range
    is_within_reading_week = term_dates['reading_week_start'] <= current_date < term_dates['reading_week_start'] + timedelta(days=5)
    
    # Check if current_date is in the unavailable dates
    is_unavailable_date = current_date in unavailable_dates
    
    # Only consider the lecture if the current_date is not within the reading week and is not an unavailable date
    if not (is_within_reading_week or is_unavailable_date):
        for lecture in data['Class section - Lecture']:
            if current_date.strftime('%A').upper() == lecture['day_of_week']:
                total_lecture_times += 1
                
    current_date += timedelta(days=1)

# Summary and Warnings
print("\nScheduling Summary:")

# Create a new table named scheduleSummaryTable
scheduleSummaryTable = PrettyTable()
scheduleSummaryTable.field_names = ["Activity scheduled", "Scheduled", "Available"]
# Add rows to the scheduleSummaryTable
scheduleSummaryTable.add_row(["Lectures", lecture_topic_index,len(lecture_topics)])
scheduleSummaryTable.add_row(["Tutorials (/§)", tutorial_topic_index, len(tutorial_topics)])
scheduleSummaryTable.add_row(["Labs (/§)", lab_topic_index, len(lab_topics)])
# Print the scheduleSummaryTable
print(scheduleSummaryTable)


# Check for unused lecture times
if total_lecture_times > len(lecture_topics):
    print_warning(f"There are {total_lecture_times - len(lecture_topics)} lecture times remaining in the term that have not been utilized with topics.")
elif total_lecture_times == len(lecture_topics):
    print("✅ All lecture times in the term have been successfully utilized with topics.")
else:
    print_warning(f"More topics provided ({len(lecture_topics)}) than available lecture times ({total_lecture_times}).")

# Check for discrepancies in scheduled lectures vs. available topics
if lecture_topic_index < len(lecture_topics):
    print_warning(f"Not all lecture topics were scheduled. {len(lecture_topics) - lecture_topic_index} topics remain unscheduled.")
elif lecture_topic_index > len(lecture_topics):
    print_warning(f"More lectures were scheduled than available topics. {lecture_topic_index - len(lecture_topics)} extra lectures were scheduled without topics.")


# Displaying scheduled activities using PrettyTable
def display_activities(activities):
    # Grouping activities by type
    activity_groups = {}
    for activity in activities:
        current_date, _, _, _, entry_type, topic = activity
        if entry_type not in activity_groups:
            activity_groups[entry_type] = []
        activity_groups[entry_type].append((topic, current_date))
    
    # Displaying each group in a PrettyTable
    for activity_type, group in activity_groups.items():
        table = PrettyTable()
        table.field_names = ["Topic", "Date"]
        table.align["Topic"] = "l"
        for entry in group:
            table.add_row(list(entry))
        print(f"\n{activity_type}\n{table}")

# Invoke the function to display activities
display_activities(scheduled_activities)

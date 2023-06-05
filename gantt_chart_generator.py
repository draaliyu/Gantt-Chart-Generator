import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import matplotlib.dates as mdates
from docx import Document
from docx.shared import Pt
import random
from matplotlib.ticker import FixedLocator, FixedFormatter

# define task names and durations
tasks = ['task1', 'task2', 'task3', 'task4']
durations = [30, 60, 45, 90]  # in days

# Request user to specify the start dates
start_dates = []
for task in tasks:
    year = int(input(f"Enter start year for {task}: "))
    month = int(input(f"Enter start month for {task}: "))
    day = int(input(f"Enter start day for {task}: "))
    start_dates.append(datetime(year, month, day))

# function to generate random colors
def generate_random_color():
    r = lambda: random.randint(0,255)
    return '#%02X%02X%02X' % (r(),r(),r())

# Create a GANTT chart
fig, ax = plt.subplots()

# Store start and end dates for x-axis
x_dates = []

for i in range(len(tasks)):
    start_date = start_dates[i]
    end_date = start_date + timedelta(days=durations[i])
    ax.barh(tasks[i], end_date - start_date, left=start_date, height=0.5, color=generate_random_color(), align='center')

    # Add a vertical line to mark the end date
    ax.axvline(x=end_date, linestyle='dashed', color='r')

    # Add start and end dates as text at the beginning and end of each bar
    ax.text(start_date, i, start_date.strftime('%b %d, %Y'), va='top', ha='right', fontsize=8)
    ax.text(end_date, i, end_date.strftime('%b %d, %Y'), va='top', ha='left', fontsize=8)

    # Add start and end dates to list
    x_dates += [start_date, end_date]

# Set the x-axis and y-axis
ax.set_xlabel('Date')
ax.set_ylabel('Tasks')

# Set the chart title
ax.set_title('Gantt Chart')

# Set x-axis labels and ticks
ax.xaxis.set_major_locator(FixedLocator(mdates.date2num(x_dates)))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d, %Y'))

# Set the x-axis limits
start_date = min(start_dates)
end_date = max(start_date + timedelta(days=d) for start_date, d in zip(start_dates, durations))
ax.set_xlim(start_date, end_date)

# Adjust the y-axis ticks and labels
ax.set_yticks(range(len(tasks)))
ax.set_yticklabels(tasks)

# Rotate the x-axis tick labels for better visibility
plt.xticks(rotation=45)

# Save chart as a png file
plt.savefig('gantt_chart.png', dpi=300, bbox_inches='tight')

# Display the chart
plt.show()

# Add the chart to a word document
document = Document()
document.add_picture('gantt_chart.png', width=Pt(500))  # Set picture width to fit the document

# Save the Word document
document.save('gantt_chart.docx')

# Gantt-Chart-Generator
This Python script was developed by Dr. Aliyu Abubakar. It allows you to generate a Gantt chart from a list of tasks, their start dates, and their durations. The resulting chart is saved as a PNG image and inserted into a Word document.

The script requires the following Python libraries:

    matplotlib
    datetime
    random
    python-docx

If you don't have these installed, you can use pip:
  pip install matplotlib python-docx

# Usage

The script asks the user for start dates (day, month, year) for each task defined in the script. It then generates a Gantt chart where each task is represented by a horizontal bar, with start and end dates labelled.

The Gantt chart is saved as a PNG image and then inserted into a Word document.

In the Word document, the tasks are listed along with their start dates and durations.

# Customization

You can customize the script to suit your needs. For example, you can add or remove tasks, or adjust the durations. The color of the bars in the Gantt chart is generated randomly for each task.

# Limitations

The script currently does not support specifying different progress rates or dependencies between tasks.

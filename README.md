# Todoist Task Manager with Excel Export

This Python script interacts with the Todoist API to fetch projects, tasks, and subtasks. It then organizes and flattens the task hierarchy into a format suitable for Excel export, allowing easy tracking of tasks and subtasks across different projects.

## Features

- **Fetch Projects**: Retrieves all the projects from your Todoist account.
- **Fetch Tasks & Subtasks**: Fetches tasks and recursively fetches all subtasks for a given project.
- **Flatten Task Hierarchy**: Dynamically flattens the task and subtask structure into a table format for exporting.
- **Export to Excel**: Updates or creates an Excel file (`.xlsx`) that lists all tasks and subtasks in a hierarchical structure.
- **Duplicate Handling**: Avoids adding duplicate tasks when updating the Excel file.

## Requirements

- Python 3.x
- Required Python libraries:
  - `requests`
  - `pandas`
  - `openpyxl` (for Excel export)

You can install the required libraries using:
```bash
pip install requests pandas openpyxl

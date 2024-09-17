import os
import requests
import json
import pandas as pd

# Define your Todoist API token
API_TOKEN = 'todoist api'

# Base URL for Todoist API
BASE_URL = 'https://api.todoist.com/rest/v2'

# Function to fetch projects
def get_projects():
    url = f'{BASE_URL}/projects'
    headers = {
        'Authorization': f'Bearer {API_TOKEN}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to fetch tasks for a given project ID
def get_tasks(project_id=None):
    url = f'{BASE_URL}/tasks'
    headers = {
        'Authorization': f'Bearer {API_TOKEN}'
    }
    params = {}
    if project_id:
        params['project_id'] = project_id
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json()

# Recursive function to get subtasks for a given task ID
def get_subtasks(task_id, all_tasks):
    subtasks = [task for task in all_tasks if task['parent_id'] == task_id]
    subtasks_with_subtasks = []
    
    for subtask in subtasks:
        subtasks_of_subtask = get_subtasks(subtask['id'], all_tasks)
        subtasks_with_subtasks.append({
            'content': subtask['content'],
            'subtasks': subtasks_of_subtask
        })
    
    return subtasks_with_subtasks

# Function to fetch all tasks (including subtasks) for a given project ID
def get_all_tasks(project_id):
    all_tasks = get_tasks(project_id)
    main_tasks = [task for task in all_tasks if not task['parent_id']]
    
    task_dict = {task['id']: {'content': task['content'], 'subtasks': get_subtasks(task['id'], all_tasks)} for task in main_tasks}
    
    return task_dict.values()

# Main function to get projects with their tasks and subtasks
def get_project_data():
    projects = get_projects()
    project_data = {}

    for project in projects:
        project_id = project['id']
        tasks = get_all_tasks(project_id)
        formatted_tasks = [{'content': task['content'], 'subtasks': task['subtasks']} for task in tasks]
        project_data[project['name']] = {
            'id': project_id,
            'tasks': formatted_tasks
        }

    return project_data

# Dynamic recursive flattening function for any level of subtasks
def flatten_task_hierarchy(project_name, task, level=1):
    rows = []
    current_row = [project_name] + [''] * (level - 1) + [task['content']]
    rows.append(current_row)
    
    for subtask in task['subtasks']:
        subtask_rows = flatten_task_hierarchy(project_name, subtask, level + 1)
        rows.extend(subtask_rows)

    return rows

# Function to flatten project, task, and subtask structure for Excel export
def flatten_data_for_excel(project_data):
    rows = []
    
    for project_name, project_info in project_data.items():
        for task in project_info['tasks']:
            task_rows = flatten_task_hierarchy(project_name, task, level=1)
            rows.extend(task_rows)
    
    return rows

# Function to load existing Excel data into a DataFrame
def load_existing_excel_data(excel_filename):
    if os.path.exists(excel_filename):
        return pd.read_excel(excel_filename)
    else:
        return pd.DataFrame()  # Return an empty DataFrame if the file doesn't exist

# Function to update the existing Excel file or create a new one, only adding new tasks/subtasks
def update_excel_file(flattened_data, excel_filename='todoist_data_dynamic_subtasks.xlsx'):
    # Calculate the maximum depth of tasks/subtasks
    max_depth = max(len(row) for row in flattened_data)
    padded_data = [row + [''] * (max_depth - len(row)) for row in flattened_data]

    # Create dynamic column headers
    columns = ['Project'] + [f'Task Level {i+1}' for i in range(max_depth - 1)]

    # Load existing Excel data
    existing_df = load_existing_excel_data(excel_filename)

    # Create DataFrame from the new flattened data
    new_df = pd.DataFrame(padded_data, columns=columns)

    if not existing_df.empty:
        # Check for duplicates by comparing the new data with the existing data
        combined_df = pd.concat([existing_df, new_df]).drop_duplicates().reset_index(drop=True)
    else:
        # If there is no existing data, just use the new data
        combined_df = new_df

    # Save the updated data to Excel
    combined_df.to_excel(excel_filename, index=False)
    print(f"Excel file '{excel_filename}' has been updated with new tasks/subtasks.")

# Main execution
if __name__ == '__main__':
    # Fetch project data
    project_data = get_project_data()

    # Flatten the project data for Excel export
    flattened_data = flatten_data_for_excel(project_data)

    # Update or create the Excel file, only adding new tasks/subtasks
    update_excel_file(flattened_data)

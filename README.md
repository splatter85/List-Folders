Folder Management Application
Overview
This application allows users to count, view and alphabetize folders by adding, removing, and searching for directories. It supports drag-and-drop functionality and can export folder names to text, CSV, or Excel files. The app is built using Python with Tkinter for the graphical user interface and includes features for undoing and redoing actions.

Installation
Prerequisites
Make sure you have Python 3 installed. You can download Python from python.org.

Dependencies
This project requires the following Python packages:

tkinter (included with Python)
tkinterdnd2 (for drag-and-drop support)
openpyxl (for Excel file handling)
To install the required packages, you can use pip. Run the following commands in your command prompt or terminal:

sh
Copy code
pip install tkinterdnd2
pip install openpyxl
Running the Application
Clone the Repository

If you haven’t already, clone the repository using:

sh
Copy code
git clone https://github.com/your-username/your-repository-name.git
Navigate to the Project Directory

sh
Copy code
cd your-repository-name
Run the Application

Execute the script using Python:

sh
Copy code
python your_script_name.py
Replace your_script_name.py with the actual name of your Python script file.

Usage
Add Folder: Click the "Add Folder" button to open a dialog for selecting folders. You can also drag and drop folders into the window.
Remove Folder: Select a folder from the list and click "Remove Folder" to delete it from the selection.
Clear All: Click "Clear All" to remove all selected folders.
Undo/Redo: Use the "Undo" and "Redo" buttons to revert or reapply actions.
Search: Click "Search" to open a new window where you can search for folders based on name and depth.
Generate File: Click "Generate File" to save the selected folder names to a file. Choose from text, CSV, or Excel formats.
Configuration
No additional configuration is required. The application starts with a default directory set to the user's home directory and does not require any configuration files.

Troubleshooting
No Folders Selected: Ensure you have added at least one folder before attempting to generate a file.
Error Messages: If you encounter issues, make sure all dependencies are installed correctly. Check the terminal or command prompt for error messages and ensure that Python and the required packages are up to date.

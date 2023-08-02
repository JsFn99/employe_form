# Employee Information Entry GUI

This is a simple employee information entry application built using Python and Tkinter. The application allows users to input and save employee details, such as first name, last name, gender, age, department, address, telephone number, and email. The information is then stored in a spreadsheet (employes.xlsx) using the openpyxl library.

## Requirements

- Python 3.x
- Tkinter (usually included with Python)

## How to Use

1. Clone or download the project files to your local machine.

2. Run the `employee_info_entry.py` script using Python:

   ```
   python employee_info_entry.py
   ```

3. The application window will appear, prompting you to enter the employee details.

4. Fill in the required information for the employee and click the "Enregistrer" button to save the data.

5. If an email already exists in the spreadsheet, an error message will be displayed.

6. After saving the employee's information, the entry fields will be cleared to allow input for the next employee.

## Features

- Stylish and user-friendly graphical interface.
- Entry fields aligned neatly in groups of three inside a rectangle titled "USER."
- Input validation to ensure all required fields are filled before saving.
- Prevention of duplicate emails (as matricule) in the spreadsheet.
- Successful saving notification in case of successful data entry.
- Automatic clearing of entry fields after saving an employee's information.

## Contributions

This project is open to contributions and improvements. If you have any suggestions or encounter any issues, feel free to create a pull request or open an issue.

Enjoy using the Employee Information Entry GUI!

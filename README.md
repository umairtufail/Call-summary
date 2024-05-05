Monthly Call Report Generator 📊📅

This Python script automates the generation of a monthly call report from an Excel dataset. Key functionalities include:

Data Loading: Utilizes Pandas to load data from an Excel file selected by the user.
Data Processing: Extracts relevant information such as caller numbers, dates, and call types from the dataset.
Data Filtering: Filters data for business hours (9 AM to 5 PM) and specific call types (Mailbox Calls, Lost Calls, Accepted Calls).
Data Grouping: Groups data by call type and month to facilitate analysis.
Visualization: Generates stacked bar plots illustrating the number of calls over months, as well as individual bar plots for each month's call types.
PDF Export: Saves the generated plots to a PDF file named "monthly_call_report.pdf".
How to Use:

Run the script and select an Excel file containing call data.
Review the generated PDF report for insights into call patterns.
Dependencies:

Pandas: pip install pandas
Matplotlib: pip install matplotlib
NumPy: pip install numpy
Tkinter: Included in standard Python library.
Note: Ensure that the Excel file follows the specified format and contains relevant call data.

Contributions and Feedback:
Feedback and contributions are welcome to enhance the functionality and usability of this script.

Disclaimer:
This script is provided as-is for educational and informational purposes. Use it responsibly and ensure compliance with data privacy regulations.

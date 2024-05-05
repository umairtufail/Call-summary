# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
import tkinter as tk
from tkinter import filedialog

# Ask user to select the Excel file
root = tk.Tk()
root.withdraw()  # Hide the main window
file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

if not file_path:
    print("No Excel file selected. Exiting.")
    exit()

# Load data from the selected Excel file
df = pd.read_excel(file_path, header=1, names=['Data'])

# Extract relevant information from the data
df[['ToNumber', 'FromNumber', 'Date', 'Type', 'Note']] = df['Data'].str.split(',', expand=True)

# Convert date strings to datetime objects
df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d %H:%M:%S %z', errors='coerce')

# Remove rows with missing dates
df = df[df['Date'].notnull()]

# Convert call types to integers
df['Type'] = df['Type'].astype(int)

# Filter data for business hours
business_hours_mask = (df.Date.dt.hour >= 9) & (df.Date.dt.hour <= 17)
df = df[business_hours_mask]

# Filter data for specific call types
df = df[df['Type'].isin([1, 2, 4])]
df['Type'] = df['Type'].replace({1: 'Mailbox Calls', 2: 'Lost Calls', 4: 'Accepted Calls'})

# Group data by call type and month
grouped_data = df.groupby([df['Type'], df['Date'].dt.to_period('M')]).size().reset_index(name='Count')

# Create a PDF file to save the plots
pdf_filename = 'monthly_call_report.pdf'
with PdfPages(pdf_filename) as pdf:
    # Generate stacked bar plot for call types over months
    ax = grouped_data.pivot(index='Date', columns='Type', values='Count').plot(kind='bar', stacked=True)
    max_y = np.ceil(ax.get_ylim()[1])
    plt.yticks(np.arange(0, max_y + 1, 1))
    plt.title('Monthly Call Report')
    plt.xlabel('Month')
    plt.ylabel('Number of Calls')
    pdf.savefig()  # Save the current plot to the PDF file
    plt.close()

    # Generate individual bar plots for each month's call types
    for month, data in grouped_data.groupby('Date'):
        ax = data.pivot(index='Date', columns='Type', values='Count').plot(kind='bar', color=['red', 'green', 'blue'])
        max_y = np.ceil(ax.get_ylim()[1])
        plt.yticks(np.arange(0, max_y + 1, 1))
        plt.title(f'Monthly Call Report for {month}')
        plt.xlabel('Call Type')
        plt.ylabel('Number of Calls')
        pdf.savefig()  # Save the current plot to the PDF file
        plt.close()

# Display a message indicating successful completion
print(f"Plots saved in '{pdf_filename}'")

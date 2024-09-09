import pandas as pd
import requests
from tkinter import Tk, filedialog

# Function to open a file dialog to select the Excel file
def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

# Function to open a file dialog to choose where to save the file
def save_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    save_path = filedialog.asksaveasfilename(title="Save As", defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    return save_path

# Select the file and read the Excel file
file_path = select_file()
if not file_path:
    print("No file selected. Exiting...")
    exit()

df = pd.read_excel(file_path)

# Create a list to store the NPIs
npi_list = []

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    first_name = row['HCP FIRST NAME']  # Replace with your actual column name
    last_name = row['HCP LAST NAME']    # Replace with your actual column name
    postal_code = row['ZIP CODE']       # Replace with your actual column name
    
    # Prepare the API request URL
    url = f"https://npiregistry.cms.hhs.gov/api/?first_name={first_name}&last_name={last_name}&postal_code={postal_code}&version=2.1"
    
    # Make the request to the API
    response = requests.get(url)
    
    if response.status_code == 200:
        data = response.json()
        if 'results' in data and data['results']:
            npi = data['results'][0]['number']  # Get the first NPI from the results
            npi_list.append(npi)
        else:
            npi_list.append('No NPI Found')
    else:
        npi_list.append('API Error')

# Add the NPI results to the DataFrame
df['NPI2'] = npi_list

# Create a new DataFrame with only the required columns
output_df = df[['HCP FIRST NAME', 'HCP LAST NAME', 'ZIP CODE', 'NPI']]

# Ask the user where to save the updated file and in which format
output_path = save_file()
if not output_path:
    print("No save location selected. Exiting...")
    exit()

# Save the new DataFrame to the selected location in the chosen format
if output_path.endswith('.xlsx'):
    output_df.to_excel(output_path, index=False)
elif output_path.endswith('.csv'):
    output_df.to_csv(output_path, index=False)

print(f"Results saved to {output_path}")

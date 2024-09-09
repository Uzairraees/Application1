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

npi_df = pd.read_excel(file_path)

# Create lists to store the results
taxonomy_descriptions = []
taxonomy_codes = []
sln_list = []

# Loop through each NPI in the DataFrame
for npi in npi_df['NPI']:  # Replace 'NPI' with the actual column name if different
    url = f"https://npiregistry.cms.hhs.gov/api/?number={npi}&version=2.1"
    response = requests.get(url)
    
    if response.status_code == 200:
        data = response.json()
        if 'results' in data and data['results']:
            # Initialize variables
            primary_taxonomy_desc = None
            primary_taxonomy_code = None
            state_license_number = None
            
            # Loop through the taxonomies to find the one with 'Primary' set to True
            for taxonomy in data['results'][0]['taxonomies']:
                if taxonomy.get('primary', False):  # Check if it's the primary taxonomy
                    primary_taxonomy_desc = taxonomy['desc']
                    primary_taxonomy_code = taxonomy['code']
                    
                    # Get state and license number
                    state = taxonomy.get('state')
                    license_number = taxonomy.get('license')
                    if state and license_number:
                        state_license_number = f"{state}-{license_number}"
                    break  # Exit the loop once the primary taxonomy is found
            
            # Handle results for each field
            taxonomy_descriptions.append(primary_taxonomy_desc if primary_taxonomy_desc else 'No Primary Taxonomy Found')
            taxonomy_codes.append(primary_taxonomy_code if primary_taxonomy_code else 'No Code Found')
            sln_list.append(state_license_number if state_license_number else 'No SLN Found')
        else:
            taxonomy_descriptions.append('No Taxonomy Found')
            taxonomy_codes.append('No Code Found')
            sln_list.append('No SLN Found')
    else:
        taxonomy_descriptions.append('API Error')
        taxonomy_codes.append('API Error')
        sln_list.append('API Error')

# Add the results to the DataFrame
npi_df['Taxonomy'] = taxonomy_descriptions
npi_df['Taxonomy Code'] = taxonomy_codes
npi_df['SLN'] = sln_list

# Create a new DataFrame with only the NPI column and the three new columns
output_df = npi_df[['NPI', 'Taxonomy', 'Taxonomy Code', 'SLN']]

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


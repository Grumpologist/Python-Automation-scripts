import os
import pandas as pd
from datetime import datetime

# Define the output folder path
output_folder_path = r"C:\Users\ocallagz\Desktop\PythonScriptTesting"
# The rest of the script continues here...
def process_data(input_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)
    
    # Generate a timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Ensure the output folder path exists
    os.makedirs(output_folder_path, exist_ok=True)

    # Create the output file name
    output_file_path = os.path.join(output_folder_path, f"ProcessedData_{timestamp}.csv")

    # Export DataFrame to a CSV file
    df.to_csv(output_file_path, index=False)
    
    print(f"Data processed successfully and saved to {output_file_path}")

# Example usage (the actual file path should be provided)
# process_data("example_input.xlsx")
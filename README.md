# MTR.py
This Python script allow to process multiple .MTR files, as well as plotting some data. Users can interactively select two points on each curve to calculate the slope. Results are saved incrementally to an .xlsx file in an output folder.

- Visual Studio Code release used: 1.93.1
- Python release used: 3.12.4. 64-bit
> Warning! There are no guaranties this code will run on your machine.

## Features
- Batch processing: script processes all .MTR files in the specified input folder, handling multiple files in one run.
- Data extraction: script reads each .MTR file, skipping the first 20 lines, and extracts data from the specified columns.
- Interactive plotting: script plots data using matplotlib and allows users to select two points on each curve to calculate the slope interactively.
- Handling slope calculation: calculates the slope between two user-selected points on the plotted curve.
- Incremental data saving: script writes the calculates slope values and corresponding filenames to an .xlsx file in the output folder after each calculation.
- Auto-handling of file creation: script creates the slope_results.xlsx file with appropriate headers if it doesn’t exist.

## Dependencies
- os: for handling file paths and directory operations.
- pandas: dor reading .MTR files and writing data to the .xlsx file.
- matplotlib: for plotting graphs and allowing interactive point selection.
- openpyxl: used by pandas for Excel file manipulation (openpyxl is the default engine used for writing .xlsx files).

## Parameters
- input_folder (hardcoded path): is a directory containing the .MTR files to process.
- output_folder (hardcoded path): is a directory where the slope_results.xlsx file will be saved.
- .MTR file structure: script sssumes the file is a comma-separated text file, with data starting after the first 20 lines. The script expects the second column as x and the third column as y.

## Description
Following script consist of several principle blocks of the code, which are explained below.

### Library imports
- os: used for file and directory operations, such as checking the existence of folders and iterating through files.
- pandas (pd): for handling data reading and writing, specifically for reading .MTR files and saving results to an Excel file.
- matplotlib.pyplot (plt): for plotting graphs and enabling interactive point selection on the plotted curve.
 ```
import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
 ```

### Defining the process_files function
- Function definition: process_files takes input_folder and output_folder as input parameters.
- Setting the output path: creates the full path for slope_results.xlsx in the output folder.
- Initialize Excel file: checks if slope_results.xlsx exists: if it doesn’t, block creates a new file with Filename and Slope columns using pandas.
 ```
def process_files(input_folder, output_folder):
    output_path = os.path.join(output_folder, 'slope_results.xlsx')

    if not os.path.exists(output_path):
        slope_df = pd.DataFrame(columns=['Filename', 'Slope'])
        slope_df.to_excel(output_path, index=False)
 ```

### .MTR datafile iteration
- File iteration: block loops through each file in input_folder and processes files with the .MTR extension.
- File reading: block reads the .MTR file using pandas.read_csv, skipping the first 20 lines which contain experiment setup data.
- Extracting data: block assumes columns' x values are in the second column (index 1), and y values are in the third column (index 2), storing them as NumPy arrays.
 ```
    for filename in os.listdir(input_folder):
        if filename.endswith(".MTR"):
            filepath = os.path.join(input_folder, filename)
            
            data = pd.read_csv(filepath, skiprows=20, header=None, delimiter=",", engine='python')

            x = data.iloc[:, 1].values
            y = data.iloc[:, 2].values
 ```

### Plotting data and allowing interactive point selection
- Plotting: block uses matplotlib to create x-y plot with a figure size of 9x9 inches.
- Customization: block adds labels, a title, a legend, and grid lines for better visualization.
- Axis limits: block sets x and y limits to match the data range, ensuring the full curve is visible.
- Interactive point selection: block uses plt.ginput(2, timeout=-1) to allow the user to select exactly two points on the plot.
- Closing the plot: block closes the plot window after point selection using plt.close().
 ```
            plt.figure(figsize=(9, 9))
            plt.plot(x, y, label='Original Data', color='blue')
            plt.title(f"Select two points on the curve for file: {filename}")
            plt.xlabel('X')
            plt.ylabel('Y')
            plt.legend()
            plt.grid(True)

            plt.xlim(min(x), max(x))
            plt.ylim(min(y), max(y))

            points = plt.ginput(2, timeout=-1)
            plt.close()
 ```

### Calculating slope and writing data to Excel
- Point validation: block checks if exactly two points were selected. If not, it skips to the next file.
- Slope calculation: block calculates the slope using the formula slope = ((y2 - y1) / (x2 - x1)).
- Data entry creation: block creates a pandas dataframe with the filename and calculated slope.
- Writing to Excel: block opens slope_results.xlsx in append mode (mode='a') and writes the new slope entry without overwriting existing data. It adds rows using startrow.
- Feedback: Prints a confirmation message for each processed file, showing the slope or indicating if selection failed.
 ```
            if len(points) == 2:
                (x1, y1), (x2, y2) = points

                slope = (y2 - y1) / (x2 - x1)

                slope_entry = pd.DataFrame([[filename, slope]], columns=['Filename', 'Slope'])

                with pd.ExcelWriter(output_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                    slope_entry.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

                print(f"Processed file: {filename}, Slope: {slope}")
            else:
                print(f"Failed to select two points for file: {filename}. Skipping.")
 ```

### Main execution block
- Main guard: block ensures the script runs only when executed directly, not when imported as a module.
- Setting paths: make sure to hardcode your own path to both input and output folders.
- Path validation: block checks if both input and output folders exist before calling process_files.
- Error handling: block displays a message if the directories don’t exist.
 ```
if __name__ == "__main__":
    input_folder = "/Users/path/to/input/folder"
    output_folder = "/Users/path/to/output/folder"

    if os.path.exists(input_folder) and os.path.exists(output_folder):
        process_files(input_folder, output_folder)
    else:
        print("Please make sure the input and output folders exist.")
 ```

## License
This project is licensed under the MIT License. See LICENSE file.

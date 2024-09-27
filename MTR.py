import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def process_files(input_folder, output_folder):
    output_path = os.path.join(output_folder, 'slope_results.xlsx')

    if not os.path.exists(output_path):
        slope_df = pd.DataFrame(columns=['Filename', 'Slope'])
        slope_df.to_excel(output_path, index=False)

    # Iterate through all .MTR files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".MTR"):
            filepath = os.path.join(input_folder, filename)

            data = pd.read_csv(filepath, skiprows=20, header=None, delimiter=",", engine='python')

            x = data.iloc[:, 1].values
            y = data.iloc[:, 2].values

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

            if len(points) == 2:
                (x1, y1), (x2, y2) = points

                slope = (y2 - y1) / (x2 - x1)

                slope_entry = pd.DataFrame([[filename, slope]], columns=['Filename', 'Slope'])

                with pd.ExcelWriter(output_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                    slope_entry.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

                print(f"Processed file: {filename}, Slope: {slope}")
            else:
                print(f"Failed to select two points for file: {filename}. Skipping.")

    print(f"All results have been saved to {output_path}")

if __name__ == "__main__":
    input_folder = "/Users/path/to/input/folder"
    output_folder = "/Users/path/to/output/folder"

    if os.path.exists(input_folder) and os.path.exists(output_folder):
        process_files(input_folder, output_folder)
    else:
        print("Please make sure the input and output folders exist.")
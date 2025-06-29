# Excel Data Corrector & Chart Generator

![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![OpenPyXL](https://img.shields.io/badge/Library-OpenPyXL-green.svg)

## Project Overview

This is a simple yet practical Python script designed to automate the process of correcting data within an Excel spreadsheet and then visualizing the adjusted data using a bar chart. It's a foundational project that demonstrates basic file manipulation, data processing, and integration with popular libraries for common office tasks.

## Problem Solved

Manually adjusting values in large Excel files can be tedious and prone to errors. This script automates a common scenario: applying a consistent correction (e.g., a discount or tax adjustment) to a column of prices and then generating a visual representation of these new values directly within the spreadsheet.

## Features

-   **Automated Price Correction:** Reads existing prices from a specified column and applies a fixed correction factor (e.g., a 10% discount) to each value.
-   **New Column for Corrected Prices:** Adds the calculated corrected prices to a new, dedicated column in the Excel sheet.
-   **Bar Chart Generation:** Dynamically creates a bar chart based on the newly calculated corrected prices, embedding it directly into the Excel file for immediate visualization.
-   **User-Friendly Interface:** Prompts the user for the Excel file name and sheet name, making it adaptable to different files.
-   **Error Handling:** Includes basic error handling to catch issues like incorrect file or sheet names.

## How It Works

The script leverages the `openpyxl` library to interact with Excel files.
1.  It loads the specified Excel workbook and sheet.
2.  It iterates through rows, accesses the value in the third column (index 3, assuming prices are here), and calculates a `corrected_price` by multiplying it by `0.9` (representing a 10% reduction).
3.  The `corrected_price` is then written to the fourth column (index 4) of the same row.
4.  Finally, it creates a `BarChart` using the `Reference` object to select the newly added corrected prices, adds the data to the chart, and places the chart on the sheet.
5.  The modified workbook is saved as `updated-[original_filename]`.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/meez-111/Python Excel Data Corrector.git
    cd Python Excel Data Corrector
    ```
2.  **Install dependencies:**
    This project requires `openpyxl`. You can install it using pip:
    ```bash
    py -m pip install openpyxl
    ```

## Usage

1.  **Prepare your Excel file:** Ensure you have an Excel file (`.xlsx`) with numerical data in the third column that you wish to correct.
2.  **Run the script:**
    ```bash
    py app.py
    ```
3.  **Follow the prompts:**
    The script will ask you to:
    * `enter the name of the exel file>` (e.g., `your_data.xlsx`)
    * `enter the name of the exel sheet>` (e.g., `Sheet1`)
4.  **Check the output:** A new Excel file named `updated-[your_file_name].xlsx` will be created in the same directory, containing the corrected prices in the fourth column and a bar chart visualizing them.

## Example

Let's say your `your_data.xlsx` `Sheet1` looks like this:

| ID | Item | Price |
| -- | ---- | ----- |
| 1  | A    | 100   |
| 2  | B    | 200   |
| 3  | C    | 150   |

After running the script, `updated-your_data.xlsx` `Sheet1` will look like this, with a chart embedded at cell `E2`:

| ID | Item | Price | Corrected Price |
| -- | ---- | ----- | --------------- |
| 1  | A    | 100   | 90.0            |
| 2  | B    | 200   | 180.0           |
| 3  | C    | 150   | 135.0           |

## Future Enhancements (Ideas for future versions)

-   Allow the user to specify the input column, output column, and the correction factor.
-   Add more robust error handling for invalid data types in cells.
-   Implement different chart types based on user choice.
-   Create a simple GUI for easier interaction.
-   Process multiple sheets or multiple files.

## Contributing

Feel free to fork this repository, make improvements, and submit pull requests. Any contributions are welcome!

## Contacts

**LinkedIn:** [My LinkedIn](https://www.linkedin.com/in/moaz-sabra-3a7565330/)

**Email:** [My Email](meez.sabra.111@gmail.com)
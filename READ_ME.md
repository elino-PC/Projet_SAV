
# Excel-Python Integration for Automated Data Processing

## Overview

This project integrates Python scripts with Excel macros to automate the generation of production reports for different solar installations. The Python scripts perform various analyses and manipulations on the data, which are then used for generating reports or further calculations.

## Features

- **Drop-Down Menu Integration:** Users can select an option from a drop-down menu in Excel, which is then passed as an argument to a Python script.
- **Dynamic Python Script Execution:** The Python script is executed with arguments dynamically obtained from Excel, allowing for flexible and customizable data analysis.
- **Cross-Platform Compatibility:** The setup is designed to run on different machines, ensuring that the solution is adaptable and easy to deploy across various environments.

## Requirements

- **Python**: Ensure that Python is installed on your machine. The project assumes the Python executable is accessible from the command line. Please view Guideline to assist in installing Python on your machine.
- **Excel**: Microsoft Excel with macro support enabled.
- **Python Libraries**: 
  - `pandas` (for data manipulation)
  - `openpyxl` (for working with Excel files)
  - `matplotlib` (for generating plots)


## Code Structure

1. **Organisation by class**
The code is structured around classes, like SolarInstallationVictron, which manage tasks related to Victron Energy installations. This makes the code easy to extend and maintain.

2. **Organisation by files**
Seperation of different type of tasks by files. For example, data collection is done in one file, and report generation in another. This way, once the API for Fronius, SMA and MC are available, only data collection will have to be coded using a new class, and the same file will be used to generate the report.

3. **Seperation of different tasks as functions**
Each function handles a specific task, such as collecting data, detecting anomalies, or generating reports. This modularity makes the code easier to understand and update.

4. **Regrouping big steps**
Complex processes are broken into smaller steps and grouped logically (such as the actual generate_report function in the report generator, it is short and easy to read). This makes it easier to follow the workflow and modify individual parts without affecting the rest of the code.


## Usage

1. Open the Excel file.
2. Select a value from the drop-down menu.
3. Click the button to trigger the Python script.
4. The script will process the data based on the selected argument and return the results as needed.

## Troubleshooting

- **Path Issues:** If the script is not running, ensure there are no issues with the file path, especially if there are spaces. Use quotation marks around paths in the macro.
- **Macro Security:** Ensure that macros are enabled in Excel, as the macro will not run if they are disabled.
- **Dependencies:** If the Python script fails to run, make sure all required Python libraries are installed.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

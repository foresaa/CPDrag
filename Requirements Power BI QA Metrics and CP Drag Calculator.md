# Requirements for Power BI QA Metrics and CP Drag Calculator

This document outlines the requirements for setting up and running the **Power BI QA Metrics and CP Drag Calculator** solution. It includes details about the required Python interpreter, libraries, Power BI, MS Project, and the specific custom fields needed for the CP Drag Calculator.

---

## **1. Python Interpreter and Libraries**

The solution uses Python to extract data from Microsoft Project files using the COM interface. Ensure the following requirements are met:

### **Python Interpreter**

- **Version**: Python 3.8 or later
- **Environment**: Any Python environment that supports the required libraries. Virtual environments like `venv` or `conda` are recommended.

### **Required Libraries**

Install the following libraries using `pip`:

| Library     | Version             | Description                                     |
| ----------- | ------------------- | ----------------------------------------------- |
| `pandas`    | Latest              | For handling and transforming data.             |
| `pywin32`   | Latest              | For interacting with Microsoft Project via COM. |
| `pythoncom` | Included in pywin32 | For managing COM threading.                     |
| `os`        | Built-in            | For file path operations.                       |
| `traceback` | Built-in            | For error reporting and debugging.              |
| `time`      | Built-in            | For adding delays in automation scripts.        |

#### **Installation Command**

To install the libraries, use:

```bash
pip install pandas pywin32
```

### 2. Power BI

The solution includes .pbix and .pbit files for visualizing QA metrics and CP Drag analysis.

Power BI Requirements
Version: Power BI Desktop (latest version recommended).
Features:
Support for custom DAX measures.
Ability to load Python scripts for data extraction.
Interactive dashboards with cards, tables, and visuals.
**Setup**
Install Power BI Desktop from the official Power BI website.
Ensure Python is installed and configured in Power BI:
Open Power BI Desktop.
`Go to File > Options and settings > Options > Python scripting.`
Set the Python home directory to the path of your Python interpreter.

# 3. Microsoft Project

The solution extracts data directly from Microsoft Project files using the COM API.

MS Project Requirements
`       Version: Microsoft Project Profesional 2016 or later (with support for VBA and COM automation).
File   Format : .mpp
Custom Fields : The CP Drag Calculator relies on specific custom fields in Microsoft Project to store and calculate critical path drag data.`

### 4. Custom Fields for the CP Drag Calculator

The following custom fields must be configured in Microsoft Project for the CP Drag Calculator to function effectively:

```
Required Custom Fields
Custom Field Name	Type	Description
Text1	Text	Phase name.
Text2	Text	Workstream name.
Number19	Number	CP Drag Elapsed Days.
Number20	Number	CP Drag Working Days.
Cost1	Cost	CP Drag Benefit (monetary value).
Text5	Text	Driving parallel task ID.
Text6	Text	Task ancestors (upstream tasks).
Text7	Text	All parallel task IDs.
```

**Field Configuration**

Open your project in Microsoft Project.
`Navigate to Project > Custom Fields.`
Assign the above names and data types to the corresponding fields.
Ensure the fields are populated with the appropriate data for each task.

### 5. Summary

By ensuring the above requirements are met, you’ll be able to:

Extract data directly from Microsoft Project files using Python scripts.
Load the data into Power BI for QA metrics and CP Drag analysis.
Visualize key project metrics through interactive dashboards in Power BI.
For any issues or questions, feel free to open an issue in the repository or refer to the documentation.

![Andel Projects Limited](perpop.png)

# Critical Path Drag and Reverse Drag Calculator for Microsoft Project

## Overview

**Critical Path Drag** is a project management concept that quantifies the extent to which each critical task delays the project's overall finish date. It provides insights into how much each critical activity contributes to the project's total duration.

**Reverse Drag**, on the other hand, identifies scenarios where shortening the duration of a critical task inadvertently increases the project's finish date. This phenomenon typically arises in complex schedules with dependencies such as Finish-to-Finish (FF), Start-to-Start (SS), or Start-to-Finish (SF) relationships.

This repository provides a VBA script to calculate **Critical Path Drag** and **Reverse Drag** for tasks in a Microsoft Project file. The script accounts for all dependency types (`Finish-to-Start`, `Start-to-Start`, `Finish-to-Finish`, and `Start-to-Finish`) and stores the calculated values in custom fields for analysis.

---

## Features

- Calculates **Critical Path Drag** for each critical task, considering overlapping tasks and dependencies.
- Detects **Reverse Drag** by simulating duration reductions and measuring project finish date impacts.
- Accounts for all dependency types: `FS`, `SS`, `FF`, and `SF`.
- Stores results in custom fields:
  - `Number20`: Critical Path Drag.
  - `Number19' : Reverse Drag.
  - `Cost1`: Drag Reduction Benefit (monetary value).
- Restores original task durations and project integrity after calculations.

---

## How the Code Works

1. **Critical Path Drag Calculation**:

   - For each critical task, the drag value is initialized as the task duration (converted to days).
   - The script evaluates overlapping non-critical tasks and adjusts the drag value based on their total float.

2. **Reverse Drag Detection**:

   - The script simulates reducing the task duration incrementally.
   - If shortening a task increases the project finish date, the reverse drag is calculated.

3. **Custom Field Outputs**:

   - Drag and Reverse Drag values are stored in the custom fields `Number20` and `Number19`, respectively.
   - The monetary value of reducing drag is stored in the `Cost1` field, based on the task's cost per day.

4. **Dependency Handling**:

   - The script processes all dependency types (`FS`, `SS`, `FF`, and `SF`) using the `TaskDependencies` collection.

5. **Restoration of Original Values**:
   - After testing for reverse drag, the script restores each task's original duration and recalculates the project to ensure accuracy.

---

## Installation in Microsoft Project

1. **Enable Developer Tab**:

   - Open Microsoft Project.
   - Go to `File > Options > Customize Ribbon`.
   - Enable the `Developer` tab.

2. **Add the VBA Script**:

   - Press `Alt + F11` to open the VBA editor.
   - In the editor, go to `Insert > Module`.
   - Copy and paste the VBA script into the module window.

3. **Save the Project File**:

   - Save the file as a macro-enabled project file (`.mppm`).

4. **Run the Script**:
   - Go to `Developer > Macros`.
   - Select `CalculateCriticalPathDragWithReverse` and click `Run`.

---

## How to Use

1. Open your Microsoft Project file.
2. Run the script as described in the installation instructions.
3. Review the results:
   - Open the `Number20` and `Number19` custom fields to view the Critical Path Drag and Reverse Drag values.
   - Open the `Cost1` field to see the Drag Reduction Benefit.

---

## Key Concepts

### Critical Path Drag

- Quantifies how much a task delays the project's end date.
- Helps identify the most impactful tasks for potential schedule optimization.

### Reverse Drag

- Identifies tasks where shortening their duration causes a delay in the overall project finish date.
- Often arises from complex dependencies like FF and SS.

### Dependency Types

- **Finish-to-Start (FS)**: The successor starts after the predecessor finishes.
- **Start-to-Start (SS)**: The successor starts concurrently with the predecessor.
- **Finish-to-Finish (FF)**: The successor finishes concurrently with the predecessor.
- **Start-to-Finish (SF)**: The predecessor finishes as the successor starts.

---

## Debugging and Logs

- During execution, the script prints debug information to the `Immediate Window` in the VBA editor, including drag and reverse drag values for each task.
- Ensure the `Immediate Window` is open (`Ctrl + G`) to view the logs.

---

## License

This script is licensed for educational and professional use. Please credit the author, Andy Forrester of Andel Projects, in any derivative works.

---

## Contribution

Feedback and contributions are welcome. Feel free to submit issues or suggestions for improving the code.

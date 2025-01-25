# Power BI QA Metrics and CP Drag Calculator Preview

<img src="perpop.png" alt="Andel Projects Limited" width="400">

This repository provides tools for enhancing **project quality assurance (QA)** and introducing **Critical Path Drag (CP Drag)** analysis for Microsoft Project (MSP). It includes two key Power BI files tailored for different use cases:

1. **QA Metrics Only (`QA_Metrics_Only.pbix`)**: A streamlined Power BI report focusing solely on QA checks for Microsoft Project schedules.
2. **QA Metrics and CP Drag Preview (`QA_Metrics_and_CP_Drag_Preview.pbix`)**: A comprehensive Power BI report combining QA checks with advanced CP Drag analysis.

---

## **Files Included**

### **1. QA Metrics Only (`QA_Metrics_Only.pbix`)**

This lightweight `.pbix` file provides essential QA checks for Microsoft Project schedules, making it an ideal starting point for users who want a focused solution for quality assurance.

#### **Features**

- Automated QA checks, such as:
  - Tasks missing dependencies.
  - Milestones with non-zero durations.
  - Resource overallocations or underallocations.
  - Tasks with negative float.
- Prebuilt visuals for:
  - QA metrics summaries.
  - Detailed insights into task and resource-level issues.

#### **Who Should Use This?**

- Project Managers or PMO staff who need a quick and efficient way to perform quality reviews without advanced MS Project expertise.

#### **How to Use**

1. Open the `QA_Metrics_Only.pbix` file in Power BI.
2. Connect the file to your Microsoft Project data.
3. Review the QA metrics and address highlighted issues.

---

### **2. QA Metrics and CP Drag Preview (`QA_Metrics_and_CP_Drag_Preview.pbix`)**

This comprehensive `.pbix` file extends the QA functionality by including metrics and visuals for critical path drag analysis. It showcases the integration of the **CP Drag Calculator Add-In [currently due to be released soon]** With Power BI for holistic project insights.

#### **Features**

- **All QA Metrics**: Includes the full suite of QA checks from the `QA_Metrics_Only.pbix` file.
- **CP Drag Analysis**:
  - Drag Working Days and Drag Elapsed Days for critical tasks.
  - Drag Cost and Drag Benefit for financial analysis.
  - Insights into driving parallel tasks and their impact on project schedules.
- **Integrated Dashboards**:
  - Combined views of QA and CP Drag metrics.
  - Drill-through capabilities to analyze issues in detail.

#### **Who Should Use This?**

- Users interested in previewing the capabilities of the CP Drag Calculator Add-In.
- Advanced project managers looking for a comprehensive solution that combines QA and CP Drag analysis.

#### **How to Use**

1. Open the `QA_Metrics_and_CP_Drag_Preview.pbix` file in Power BI.
2. Connect the file to your Microsoft Project data or use the included example dataset.
3. Explore QA and CP Drag dashboards to identify quality issues and critical path optimizations.

---

## **Key Benefits**

### **For QA Metrics**

- **Streamlined QA Process**: Simplifies quality assurance with prebuilt checks and dashboards.
- **No Advanced MS Project Knowledge Required**: Empowers junior PMs and PMO staff to perform quality reviews.
- **Scalable**: Works for single projects or multiple projects in a portfolio.

### **For CP Drag Integration**

- **Actionable Insights**: Highlights critical tasks and their impact on project delays and costs.
- **Enhanced Decision-Making**: Combines QA metrics with CP Drag data for a holistic view of project health.
- **Portfolio Analysis**: Aggregates CP Drag metrics across projects to identify trends and bottlenecks.

---

## **Getting Started**

### **1. Download the Files**

- [QA Metrics Only (`QA_Metrics_Only.pbix`)](./QA_Metrics_Only.pbix)
- [QA Metrics and CP Drag Preview (`QA_Metrics_and_CP_Drag_Preview.pbix`)](./QA_Metrics_and_CP_Drag_Preview.pbix)

### **2. Prepare Your Data**

- Ensure your MSP file is structured with key fields (e.g., task names, dependencies, start/finish dates, and durations).
- Use the `QA_Metrics_Only.pbix` file for focused QA or the `QA_Metrics_and_CP_Drag_Preview.pbix` file to explore CP Drag metrics.

### **3. Explore the Dashboards**

- Use Power BI’s interactive features to filter, drill down, and analyze your project data.
- Address QA issues and identify critical path optimizations.

---

## **Planned Enhancements**

- **Full CP Drag Add-In Integration**: Future updates will fully support CP Drag metrics exported directly from the add-in.
- **Portfolio-Level Reporting**: Enhanced visuals for analyzing QA and CP Drag metrics across multiple projects.
- **AI-Assisted Recommendations**: Planned integration of AI features to suggest actions based on QA and CP Drag data.

---

## **Feedback and Contributions**

We welcome feedback and contributions to improve these tools! If you encounter issues or have suggestions, please open an issue or submit a pull request.

---

## **About the CP Drag Calculator Add-In**

The **CP Drag Calculator Add-In** is an advanced tool for calculating and visualizing **critical path drag** directly within MS Project. It enables project managers to:

- Identify tasks contributing the most to project delays.
- Quantify the cost and time impact of critical path tasks.
- Optimize project schedules by focusing on high-drag tasks.

This add-in is currently in development, and its metrics will seamlessly integrate into Power BI for broader analysis.

---

## **License**

This project is distributed under the [MIT License](./LICENSE). Feel free to use and modify these files to suit your project needs.

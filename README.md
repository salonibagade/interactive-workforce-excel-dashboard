# 📊 Interactive Workforce & Call Center Performance Dashboard (Excel + VBA)

An end-to-end interactive workforce reporting dashboard built in Microsoft Excel. This project transforms multi-sheet operational logs into an executive-ready interface with parameter-driven controls and custom VBA automation.

---

## 🚀 Key Features & Interactivity
- **Dynamic Record Exploration:** Integrated form control spinners allowing stakeholders to scale visible rows dynamically.
- **Custom VBA UI Toggles:** Programmed Hide/Show macros paired with interactive eye icons to expand or collapse visual charts for workspace optimization.
- **Visual Analytics:** Designed paired Column and Pie charts tracking cumulative hours and active customer time distribution.

---

## 📁 Architecture & Multi-Sheet Pipeline
- **Sheet 1 (Break Analytics):** Granular tracking of break durations (Training, Meetings, Feedback, Unwanted breaks).
- **Sheet 2 (Efficiency Tracking):** Call center metrics, work efficiency percentages, active customer handling, and hold patterns.
- **Sheet 3 (Overall Performance):** Call routing logs, transfer rates, and average hold times.
- **Sheet 4 (Consolidated Summary):** Clean intermediate calculation layer aggregating multi-source operational metrics.
- **Sheet 5 (Dashboard UI):** Dynamic visual presentation layer housing the charts, spinners, and VBA controls.

---

## 🛠️ Tools & Techniques Used
- **Microsoft Excel:** Form Controls (Spinners), Dynamic Chart Ranges, Conditional Formatting.
- **VBA / Macros:** Shape visibility automation (`msoTrue` / `msoFalse`).
- **Data Structuring:** Data cleaning, aggregation, and decoupled reporting architecture.

---

## 📥 How to Run the Project
1. Download the `.xlsm` file from this repository.
2. Open in Microsoft Excel and click **Enable Macros / Enable Content** when prompted.
3. Use the spinners and eye toggles on the **Dashboard** sheet to explore the data dynamically.

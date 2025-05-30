# 🛒 Superstore Sales Dashboard (Excel)

This Excel project transforms raw fictional sales data from Kaggle's [Sample Superstore Dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) into a fully interactive, visually engaging **Sales Performance Dashboard** using Power Query, Pivot Tables, Charts, KPIs, and VBA.

---
## 📈 Tools Used

| Tool / Skill        | Description                                               |
|---------------------|-----------------------------------------------------------|
| **Power Query**     | Data cleaning and transformation                          |
| **Pivot Tables**    | Data aggregation and slicing                              |
| **Pivot Charts**    | Visual representation of performance indicators           |
| **VBA (Macros)**    | Interactivity: show/hide panels, auto-refresh pivots      |
| **Slicers**         | Seamless dynamic filtering                                |
| **Formulas**        | INDEX-MATCH, VLOOKUP, IF-AND-OR logic for KPI/analysis    |

---
## 📷 Preview
<img src = "https://github.com/user-attachments/assets/582e3405-1084-4664-ba85-20c129ccd60a" width = "1000px" height = "600px"><br>
<img src = "https://github.com/user-attachments/assets/69fa6d51-0c60-43cc-846e-cb287e556101" width = "1000px" height = "600px"><br>
---

## 📊 Project Highlights

### ✅ Raw Data to Dashboard Workflow:
- **Power Query** used to clean the dataset:
  - Handled data type mismatches.
  - Filtered and shaped raw data as required.
- Performed **Exploratory Data Analysis (EDA)** to understand sales trends and target metrics.

### 📌 Key Features:
- **Interactive Dashboard Design**:
  - Built using Pivot Charts, Slicers, and KPIs.
  - KPI Cards: Total Sales, Total Orders, Profit Margin, Customer Count, Average Order Value.
  - Trendline: Quarterly Sales Trend.
  - Comparative Chart: Actual Sales vs Target Sales.
  - Segmented Analysis: Region vs Segment, Sales and Profit by Category/Sub-Category.
  - Choropleth Map: Order Count by U.S. State.

- **User Interactivity**:
  - Slicer filters for dynamic exploration (Year, Category, Sub-Category, Region, Segment).
  - Custom-designed **Filter Panel** using **Shapes + VBA** that mimics Power BI bookmarks:
    - Toggle panel visibility with a click for a clean UI experience.

- **VBA Integration**:
  - VBA script to **toggle the slicer panel visibility**.
  - VBA script to **refresh all Pivot Tables** dynamically.
    
  ### 🔁 Refresh All Pivot Tables - VBA Macro
    This VBA macro loops through all worksheets in the workbook and refreshes every Pivot Table.

    ```vb
    Sub RefreshAllPivotTables()
        Dim ws As Worksheet
        Dim pt As PivotTable
    
        ' Loop through each worksheet in the workbook
        For Each ws In ThisWorkbook.Worksheets
            ' Check if the worksheet has any pivot tables
            If ws.PivotTables.Count > 0 Then
                ' Loop through each pivot table in the worksheet
                For Each pt In ws.PivotTables
                    pt.RefreshTable ' Refresh the pivot table
                Next pt
            End If
        Next ws
    
        MsgBox "✅ All Pivot Tables have been refreshed!", vbInformation, "Refresh Complete"
    End Sub

![image](https://github.com/user-attachments/assets/aeda7751-9f66-478b-ac64-ae8747ce4fb1)

  ### 🔁 Slicer Panel Show/Hide - VBA Macro
    This VBA macro toggle the Slicer Panel Visibility.

    ```vb
    Sub ShowSlicerPanel()
      ' Show the slicer panel group
      ActiveSheet.Shapes("SlicerPanel_main").Visible = True
      ActiveSheet.Shapes("ShowPanelBtn").Visible = False
      ActiveSheet.Shapes("HidePanelBtn").Visible = True
    End Sub

    Sub HideSlicerPanel()
        ' Hide the slicer panel group
        ActiveSheet.Shapes("SlicerPanel_main").Visible = False
        ActiveSheet.Shapes("ShowPanelBtn").Visible = True
        ActiveSheet.Shapes("HidePanelBtn").Visible = False
    End Sub
---
- **Advanced Excel Functions** (demonstrated in supporting sheets):
  - `INDEX-MATCH` for reverse lookups.
  - `VLOOKUP` for customer-region mapping.
  - `IF`, `AND`, `OR` logic for flagging key metrics.
  - `TEXT`, `ROUND`, and `CONCATENATE` for KPI formatting.

---


## 📁 Dataset

- **Name**: `Sample - Superstore.csv`
- **Source**: [Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)
- Contains fictional sales data across categories, sub-categories, regions, and order dates.

---

## 🎯 What I Learned

- How to take a raw CSV dataset and fully automate insights using Excel tools.
- Built a UI experience inside Excel that mimics modern BI tools like Power BI.
- Implemented advanced interactivity and automation using VBA scripting.
- Strengthened skills in Power Query, data modeling, KPI visualization, and dashboard storytelling.

---






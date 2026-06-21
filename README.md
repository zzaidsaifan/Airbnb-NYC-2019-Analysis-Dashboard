# Airbnb NYC 2019 Analysis Dashboard Using Excel, Pivot Tables & VBA

## Project Overview

This project analyzes the Airbnb New York City 2019 dataset using Microsoft Excel. The objective is to transform raw Airbnb listing data into meaningful business insights through Pivot Tables, interactive dashboards, and VBA automation.

The dashboard provides insights into listing distribution, pricing trends, room type performance, revenue potential, review activity, host concentration, and availability patterns across New York City boroughs.

---

## Dataset

**Dataset:** AB_NYC_2019.csv

The dataset contains Airbnb listings across New York City and includes information such as:

- Listing ID
- Host ID
- Host Name
- Neighborhood
- Borough
- Room Type
- Price
- Number of Reviews
- Reviews per Month
- Availability (365 Days)
- Host Listing Count

---

## Project Objectives

- Analyze Airbnb listing distribution across NYC boroughs.
- Compare average pricing across different locations.
- Identify the most common room types.
- Estimate revenue potential by borough.
- Evaluate review activity and customer engagement.
- Analyze host concentration and listing ownership.
- Build an interactive Excel dashboard for business reporting.
- Automate reporting tasks using VBA.

---

## Tools & Technologies

- Microsoft Excel
- Pivot Tables
- Pivot Charts
- Slicers
- Conditional Formatting
- VBA (Visual Basic for Applications)
- Dashboard Design
- Data Analysis

---

## Data Preparation

### Data Cleaning

The following preprocessing steps were performed:

- Removed duplicate records
- Checked for missing values
- Formatted currency fields
- Converted raw data into an Excel Table
- Created calculated fields for analysis

### Calculated Field

**Potential Revenue**

```excel
=Price*Availability_365
```

This metric estimates the annual revenue potential for each Airbnb listing.

---

## Analysis Performed

### 1. Listings by Borough

**Objective:** Determine which NYC borough contains the highest number of Airbnb listings.

**Pivot Table Configuration**

Rows:
- neighbourhood_group

Values:
- Count of Listing ID

**Visualization**
- Clustered Column Chart

---

### 2. Average Price by Borough

**Objective:** Compare pricing across different boroughs.

**Pivot Table Configuration**

Rows:
- neighbourhood_group

Values:
- Average of Price

**Visualization**
- Bar Chart

---

### 3. Room Type Distribution

**Objective:** Analyze the distribution of listing types.

**Pivot Table Configuration**

Rows:
- room_type

Values:
- Count of Listing ID

**Visualization**
- Pie Chart

Room Types:
- Entire Home/Apartment
- Private Room
- Shared Room

---

### 4. Revenue Potential Analysis

**Objective:** Estimate revenue opportunities across boroughs.

**Pivot Table Configuration**

Rows:
- neighbourhood_group

Values:
- Sum of Potential Revenue

**Visualization**
- Column Chart

---

### 5. Reviews Analysis

**Objective:** Evaluate customer engagement and listing popularity.

**Pivot Table Configuration**

Rows:
- neighbourhood_group

Values:
- Sum of Number of Reviews

**Visualization**
- Line Chart

---

### 6. Availability Analysis

**Objective:** Understand listing occupancy potential.

**Pivot Table Configuration**

Rows:
- neighbourhood_group

Values:
- Average Availability_365

**Visualization**
- Bar Chart

---

### 7. Top 10 Neighborhoods

**Objective:** Identify neighborhoods with the highest number of listings.

**Pivot Table Configuration**

Rows:
- neighbourhood

Values:
- Count of Listing ID

Filter:
- Top 10

**Visualization**
- Horizontal Bar Chart

---

### 8. Top Hosts Analysis

**Objective:** Identify hosts managing the highest number of listings.

**Pivot Table Configuration**

Rows:
- host_name

Values:
- Maximum Host Listing Count

Filter:
- Top 10

**Visualization**
- Column Chart

---

## Dashboard Features

### KPI Cards

The dashboard includes the following key performance indicators:

- Total Listings
- Average Price
- Total Reviews
- Total Hosts

### Interactive Filters

Slicers were implemented for:

- Borough
- Room Type
- Host Name

These slicers dynamically update all Pivot Tables and charts.

---

## VBA Automation

### Refresh Dashboard

```vb
Sub RefreshDashboard()

    ThisWorkbook.RefreshAll

    MsgBox "Dashboard Updated Successfully!"

End Sub
```

### Clear Filters

```vb
Sub ClearFilters()

    Dim sc As SlicerCache

    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc

    MsgBox "Filters Cleared!"

End Sub
```

### Export Dashboard as PDF

```vb
Sub ExportDashboard()

    Sheets("Dashboard").ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=ThisWorkbook.Path & "\Airbnb_Dashboard.pdf"

    MsgBox "Dashboard Exported!"

End Sub
```

---

## Key Insights

- Manhattan generally has the highest average Airbnb prices.
- Entire Home/Apartment listings dominate the market.
- Brooklyn and Manhattan contain the largest number of listings.
- Revenue potential varies significantly across boroughs.
- A small group of hosts manages a large number of listings.
- Availability trends provide insights into occupancy opportunities.

---

## Skills Demonstrated

- Data Cleaning
- Excel Analytics
- Pivot Tables
- Pivot Charts
- Dashboard Development
- KPI Reporting
- VBA Automation
- Business Intelligence
- Data Visualization
- Revenue Analysis

---


## Future Enhancements

- Power Query Integration
- Power Pivot Data Modeling
- Revenue Forecasting
- Occupancy Rate Analysis
- Geographic Mapping
- Power BI Dashboard Version

---

## Author

**Zaid Arif Saifan**

MSc Financial Engineering | HEC Montréal

### Skills
Excel • VBA • Financial Analysis • Data Analytics • Dashboard Development • Business Intelligence • Python

---

## Project Outcome

Successfully developed an interactive Excel dashboard that converts Airbnb NYC listing data into actionable business insights using Pivot Tables, KPI reporting, data visualization, and VBA automation.

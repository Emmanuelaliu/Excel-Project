## Excel-Project
A data cleaning, analysis and visualisation project with Microsoft Excel 

# Airbnb New York: Excel Data Analytics Capstone

An end-to-end Excel project that cleans Airbnb New York City listing data and turns it into an interactive dashboard, with findings and business recommendations.

<a href="https://drive.google.com/file/d/1IBVdnwNQ3AuJUgwpkLNCXNsfINkYNYRy/view?usp=drivesdk" target="blank" rel="noopener noreferrer">Dashboard Overview</a>

---

## Table of Contents

- [Project Overview](#project-overview)
- [Business Questions](#business-questions)
- [Dataset](#dataset)
- [Tools and Excel Skills Used](#tools-and-excel-skills-used)
- [Data Cleaning](#data-cleaning)
- [Dashboard](#dashboard)
- [Key Insights](#key-insights)
- [Recommendations](#recommendations)
- [Repository Structure](#repository-structure)
- [How to View This Project](#how-to-view-this-project)
- [Author](#author)

---

## Project Overview

Airbnb operates about **48,895 room listings** across New York City. The goal of this project is to clean the raw listing data and analyze the company's operations across **five years (2021 to 2025)**, then present the results in an interactive Excel dashboard that a business team can use to understand room supply, pricing, occupancy and host activity.

---

## Business Questions

The dashboard was built to answer these questions:

1. How many rooms were launched each year?
2. How many rooms does each host manage?
3. How many rooms are there by room type?
4. What is the average room price by room type?
5. How many rooms are occupied versus free?
6. How many rooms are there in each city (borough)?
7. What is the average room price by city?
8. How many neighborhoods are there by city?

---

## Dataset

The workbook contains two source sheets:

| Sheet | Description |
|-------|-------------|
| **Data and visualization** | Main working sheet with Room ID, Host name, Price, Launch date and Availability, plus the cleaned columns, pivot tables and dashboard |
| **Rooms** | Lookup sheet with room description, neighborhood, city, coordinates, room type and launch date |

**Key fields:** Room ID, Host_name, Description, Neighborhood, City, Latitude, Longitude, Room_type, Price (Amount), Launch_date, Availability, Launch_year, Status

**Size:** about 48,895 rows (one row per room listing)

---

## Tools and Excel Skills Used

- **Microsoft Excel**
- **Lookup formulas:** `VLOOKUP` to pull description, neighborhood, city, latitude, longitude and room type from the Rooms sheet
- **Date functions:** `YEAR` to create the launch year
- **Logical functions:** `IF` to create the occupancy status
- **Text functions:** `PROPER` to standardize description text
- **PivotTables** to summarize the data
- **Charts:** column, bar, funnel, doughnut, pie, 3-D column, line and area charts
- **Slicers** for interactive filtering by host
- **Dashboard design** to bring all visuals into a single view

---

## Data Cleaning

| Step | Task | Method |
|------|------|--------|
| 1 | Fill in Description, Neighborhood, City, Latitude, Longitude and Room type from the Rooms sheet | `VLOOKUP(Room ID, Rooms!range, column, FALSE)` |
| 2 | Create a `Launch_year` column | `=YEAR(Launch_date)` |
| 3 | Create a `Status` column | `=IF(Availability = 0, "OCCUPIED", "FREE")` |
| 4 | Standardize description text | `=PROPER(Description)` |

---

## Dashboard

The dashboard contains eight visuals built from PivotTables:

| Visual | Chart Type |
|--------|-----------|
| Number of rooms by launch year | Column chart |
| Number of rooms by host | Bar chart |
| Number of rooms by room type | Funnel chart |
| Average room price by room type | Doughnut chart |
| Number of rooms by status | Pie chart |
| Number of rooms by city | 3-D column chart |
| Average room price by city | Line chart |
| Number of neighborhoods by city | Area chart |


---

## Key Insights

**Growth over time**
- Room listings grew from **9,390 in 2021** to **10,116 in 2025**, with a dip in 2023 (9,711).
- Filtering by host showed **Elizabeth Andrea** had a noticeable drop in rooms hosted in 2023 (1,146, down from 1,209 in 2022), which contributes to the 2023 dip.

**Hosts**
- Eight hosts manage almost the same number of rooms, between about 6,070 and 6,175 each.
- **Elizabeth Andrea** manages the most rooms (6,175), while **James Tyler** (6,068) and **Laura Claudio** (6,070) manage the fewest.

**Room types**
- **Entire home/apt** (25,409) and **Private room** (22,326) dominate by volume. Shared rooms are a very small share.
- Entire homes have the highest average price (**$211.79**), followed by private rooms (**$89.78**) and shared rooms (**$70.13**).

**Occupancy**
- **64% of rooms (31,362) are free** and **36% (17,533) are occupied**.

**Cities**
- **Manhattan** (21,661) and **Brooklyn** (20,104) have the most rooms, followed by Queens (5,666), the Bronx (1,091) and Staten Island (373).
- Average price is highest in **Manhattan ($196.88)**, then Brooklyn ($124.38) and Staten Island ($114.81), Queens ($99.52), with the **Bronx lowest at $87.50**.

---

## Recommendations

1. **Focus on Entire home/apt and Private room listings.** They make up the bulk of the inventory and, for entire homes, carry the highest prices.
2. **Run more marketing campaigns to reduce the share of free rooms.** Almost two-thirds of rooms are currently free.
3. **Look into the 2023 dip** and the drop in Elizabeth Andrea's hosted rooms to understand what caused it and how to avoid it.
4. **Prioritize Manhattan and Brooklyn**, where supply and prices are highest.

---

## Repository Structure

```
airbnb-new-york-excel-analysis/
│
├── data/
│   └── EXCEL_CAPSTONE_PROJECT_ALIU_EMMANUEL.xlsx    # Workbook with data, pivots and dashboard
│
├── report/
│   └── EXCEL_CAPSTONE_PROJECT_ALIU_EMMANUEL.pdf     # Full written report
│
└── README.md
```

---

## How to View This Project

1. Click the Excel file in the `data` folder and choose **Download** (GitHub cannot display interactive Excel dashboards online).
2. Open the file in **Microsoft Excel 2016 or later**. A funnel chart is used, so older versions may not display it.
3. Go to the **Data and visualization** sheet to see the cleaned data, PivotTables and dashboard.

The full written report is available as a PDF in the `report` folder.

---

## Author

**Aliu Emmanuel**

- GitHub: [emmanuel-aliu](https://github.com/your-username)
- LinkedIn: [your-profile](https://www.linkedin.com/in/aliu-godwin-emmanuel)
- Email: your-email@aliuemmanuel31@gmail.com

Feedback and suggestions are welcome.

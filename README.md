# 📊 Volunteers Impact Dashboard — Excel Project

> **Tool:** Microsoft Excel | **Goal:** Clean, analyze, and visualize volunteer campaign data into an interactive dashboard.

---

## 🗂️ Workflow Overview

```
Raw Data (Volunteers Sheet)
    ↓
Volunteers_EDA       ← Exploratory Data Analysis & Quality Check
    ↓
Volunteers_clean     ← Cleaned Data for Analysis
    ↓
Pivot_Tables         ← Aggregated Data via Pivot Tables
    ↓
Volunteers_Impact_Dashboard  ← Final Dashboard
```

---

## 🔍 Phase 1 — Exploratory Data Analysis (EDA)

### Step 1 — Copy & Rename Sheet

| Action | Formula / Tool | Result | Reflection |
|---|---|---|---|
| Copy Volunteers sheet | Right-click → Move & Copy | Copy of the raw data | Original sheet stays untouched |
| Rename sheet | Rename | Sheet named `Volunteers_EDA` | Dedicated sheet for EDA work |

---

### Step 2 — Duplicate Detection

**Create `DuplicateCheckKey` field**
```excel
=[@Email]&[@CampaignID]&[@SignupDate]
```
> Concatenates key fields into a single non-human-readable identifier to catch duplicate entries.

**Create `DuplicateFlag` field**
```excel
=IF(COUNTIF([DuplicateCheckKey],[@DuplicateCheckKey])>1,"Duplicate","Unique")
```
> Flags each row as `Duplicate` or `Unique`.

**Highlight duplicates visually**
```
Home tab → Conditional Formatting → Highlight Cell Rules → Text That Contains → "Duplicate"
```
> Duplicate cells are highlighted in **red** for easy identification.

---

### Step 3 — Missing Email Detection

**Create `MissingEmail` field**
```excel
=IF([@Email]="","Missing","OK")
```
> Flags rows with no email address.

**Highlight missing values visually**
```
Home tab → Conditional Formatting → Highlight Cell Rules → Text That Contains → "Missing"
```
> Missing email cells are highlighted in **red**.

---

### Step 4 — Data Quality Summary Table

| Metric | Formula |
|---|---|
| Total volunteer records | `=COUNTA(VolunteerTable4[VolunteerID])` |
| No. of duplicate records | `=COUNTIF(VolunteerTable4[DuplicateFlag],"Duplicate")` |
| No. of missing emails | `=COUNTIF(VolunteerTable4[MissingEmail],"Missing")` |
| No. of unique campaigns | `=COUNTA(UNIQUE(VolunteerTable4[CampaignID]))` |

---

### Step 5 — Campaign Lookup

**Create `CampaignLookup` field**
```excel
=XLOOKUP([@CampaignID],CampaignTable[CampaignID],CampaignTable[CampaignName])
```
> Maps each `CampaignID` to its human-readable `CampaignName` from the Campaign reference table.

---

### Step 6 — Dashboard Questions

Based on the cleaned data, the following questions guide the dashboard design:

1. 🗺️ Which areas have the highest number of volunteers?
2. 🏆 Which campaign generated the most impact?
3. 💬 What is the overall sentiment towards their cause?
4. 🚫 How many volunteers failed to show up?
5. ⚠️ How many missing values and duplicates were found in the source datasets?

---

### Step 7 — Mock Dashboard Wireframe

**Tool:** `app.diagrams.net`

Planned layout:
- 3 Scorecards
- 3 Bar / Column Charts
- 1 Donut Chart

> Wireframing before building ensures a clear layout plan and avoids rework.

---

## 🧹 Phase 2 — Data Cleaning (`Volunteers_clean`)

### Step 8 — Copy & Rename Sheet

Copy the original `Volunteers` sheet and rename it `Volunteers_clean`. All cleaning is done here to preserve the original data.

---

### Step 9 — Remove Duplicates

```
Data tab → Remove Duplicates → Uncheck [SentimentScore] and [Feedback] → OK
```
> `SentimentScore` and `Feedback` are subjective fields and excluded from the duplicate criteria check.

---

### Step 10 — Create New Fields

**`EmailStatus`** — Flags missing emails
```excel
=IF([@Email]="","Missing","OK")
```

**`AttendanceFlag`** — Simplifies attendance status
```excel
=IF([@AttendanceStatus]="No show","Did Not Attend","Attend")
```

**`CampaignName`** — Enriches data with campaign names via lookup
```excel
=XLOOKUP([@CampaignID],CampaignTable[CampaignID],CampaignTable[CampaignName],"No CampaignName found for this CampaignID given")
```

---

## 📐 Phase 3 — Pivot Tables (`Pivot_Tables` Sheet)

A dedicated `Pivot_Tables` sheet is created to serve all dashboard visuals.

```
Select table from Volunteers_clean → Insert → Pivot Table → New Sheet → OK
```

| Pivot Table | Rows | Values | Chart Type | Dashboard Use |
|---|---|---|---|---|
| Volunteer Count | — | VolunteerID (Count) | Scorecard | Total active volunteers |
| Email Status | EmailStatus | EmailStatus (Count) | Scorecard | Missing emails count |
| Duplicate Count | — | *(from EDA sheet)* | Scorecard | Duplicate records count |
| Sentiment by Campaign | CampaignName | SentimentScore (Avg) | Clustered Bar Chart | Top 5 campaigns by sentiment |
| Volunteer Sentiment | SentimentScore | VolunteerID (Count) | Clustered Column Chart | Sentiment distribution |
| Attendance | AttendanceStatus | AttendanceFlag (Count) | Doughnut Chart | Attendance vs No-show |
| Volunteers by Region | Region | VolunteerID (Count) | Clustered Column Chart | Regional volunteer spread |

---

## 📊 Phase 4 — Dashboard (`Volunteers_Impact_Dashboard` Sheet)

### Scorecards

| Scorecard | Formula |
|---|---|
| Active Volunteers | `=GETPIVOTDATA("VolunteerID",'Pivot Tables'!$A$3)` |
| Missing Emails | `=GETPIVOTDATA("EmailStatus",'Pivot Tables'!$A$8,"EmailStatus","Missing")` |
| Duplicate Records | `=Volunteers_EDA!U3` |

> Each scorecard uses **Merge & Center** + an **inserted text box** for clean presentation.

---

### Charts

Each chart is built from its corresponding pivot table, then moved to the `Volunteers_Impact_Dashboard` sheet.

```
Select Pivot Table → Insert → [Chart Type] → OK → Format → Move to Dashboard Sheet
```

---

### Step 11 — Final Styling

**Borders**
```
Home tab → Outside Borders
```
> Applied to all scorecards, charts, and the dashboard title for a clean, structured look.

**Color Coding**
```
Click each chart → Fill Color: Light Green
Click key/critical value → Fill Color: Darker Green or Red
```
> Highlights values that require attention, making insights immediately visible.

---

## ✅ Final Output

A fully functional Excel dashboard containing:

- ✅ **3 Scorecards** — Total Volunteers · Missing Emails · Duplicates
- ✅ **2 Clustered Column Charts** — Volunteer Sentiment · Volunteers by Region
- ✅ **1 Clustered Bar Chart** — Top Campaigns by Sentiment
- ✅ **1 Doughnut Chart** — Attendance Rate

---

## 📁 Sheet Structure

```
📊 Volunteers                    ← Original raw data (untouched)
📊 Volunteers_EDA                ← Data quality checks & exploration
📊 Volunteers_clean              ← Cleaned data ready for analysis
📊 Pivot_Tables                  ← All pivot tables feeding the dashboard
📊 Volunteers_Impact_Dashboard   ← Final dashboard
📊 Campaign                      ← Campaign reference table
```

# 📊 About - Morning Meeting Report
**For Better Experience, Use Latest Version of MS OFFICE**  
**Contribution : @Dhamo0202**

---

### ⚠️ In case you encounter any security warnings or disabled macros:

Follow these steps to enable full functionality:

---

#### ✅ Check the Trust Center Settings:
1. Open **Word** or **Excel**.
2. Click on the **File** tab.
3. Select **Options** at the bottom of the left bar.
4. In the left pane, choose **Trust Center**, then click **Trust Center Settings**.
5. In the Trust Center dialog box, click **Macro Settings**.
6. Select **"Enable all macros"** *(recommended only for trusted files)*.
7. Click **OK** to apply.

---

#### 🔓 Unblock the File:
1. Close the workbook.
2. Right-click the file in File Explorer.
3. Select **Properties**.
4. Under the **General** tab, check the box labeled **Unblock** in the Security section.
5. Click **Apply**, then **OK**.
6. Reopen the workbook—macros should now run without restriction.

---

# 📊 Morning Meeting Report

A macro-enabled Excel workbook (`Morning_meeting_report.xlsm`) that collects daily production rejection data and automatically publishes it as a formatted PowerPoint presentation for morning meetings.

> Built with VBA · Excel · PowerPoint automation · DAMO Automation

---

## 🗂️ Workbook Structure

| Sheet | Purpose |
|-------|---------|
| `Startup` | Log startup rejections — machine, model, part, defect codes (1–10), A/B/C counts, cost, power-cut data |
| `Process` | Log in-process rejections — same structure as Startup |
| `Summary` | Auto-calculated daily totals: OK/NG counts, all rejection categories, changeover counts |
| `Publish` | Configuration — report date, PowerPoint output path, summary slide number, cost file path |
| `BackupData` | Raw data archive |

---

## 🚀 How to Use

1. **Open the file** and enable macros when prompted.
2. **Set the date** in `Publish!C4` (format: `DD.MM.YYYY`).
3. **Enter rejection data** in the `Startup` and `Process` sheets.
4. *(Optional)* Click **Get Costs** to auto-fill unit costs from `COST.xlsx`.
5. Click **Publish** to generate the PowerPoint report.

---

## 🔘 Macro Buttons

### 🚀 Publish — `UpdateMorningMeetingSlides`
The main button. Runs the full pipeline:
1. Reads the PowerPoint path from `Publish!C5`
2. Refreshes formulas and serial numbers on Startup & Process sheets
3. Opens PowerPoint and updates the date on every slide
4. Populates Slides 1–3 with Startup rejection data (12 rows per slide)
5. Populates Slide 4 with Process rejection data
6. Fills the Summary slide (slide number from `Publish!C6`) with daily totals
7. Shows a success message on completion

---

### 🗑️ Clear Data
| Button | Clears |
|--------|--------|
| `Clear_Data_Startup` | Startup sheet entry data (rows 11–100). Asks for **two confirmations**. Cannot be undone. |
| `Clear_Data_Process` | Process sheet entry data (rows 10–54). Same double-confirmation. Cannot be undone. |

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

---

### ➕ Dash Toggle
Fills empty rejection cells with ` - ` for clean PowerPoint table formatting, or removes them for re-entry.

| Button | Action |
|--------|--------|
| `AddDashToStartupColumns` | Fills blanks in `Startup!E:G` with `-` |
| `RemoveDashFromStartupColumns` | Clears `-` placeholders from `Startup!E:G` |
| `AddDashToProcessColumns` | Fills blanks in `Process!E:G` with `-` |
| `RemoveDashFromProcessColumns` | Clears `-` placeholders from `Process!E:G` |

---

### 💰 Cost Sync
| Button | Action |
|--------|--------|
| `FillCostsFromExternalFile` | Looks up Model + Part Name in `COST.xlsx` and fills cost into column T (Startup) or U (Process) |
| `UpdateCostsFromStartup` | Pushes costs from the Startup sheet back into `COST.xlsx`. Works via a temp copy — original is safe if anything fails. Adds new rows for unrecognised parts. |

---

## ⚡ Background Automations

These run automatically — no button needed.

| Trigger | Sheet | What It Does |
|---------|-------|--------------|
| `Worksheet_Change` | Startup | On any edit to cols E–W, writes a validation result to col AA. Compares A/B/C totals vs. defect-code totals vs. power-cut totals. Shows `TOTAL ERROR`, `POWER CUT ERROR!`, `NG Total and Power total both mistake`, or blank if all match. |
| `Worksheet_Change` | Process | On any edit to cols E–W, writes `NG ERROR` or blank to col V by comparing A/B/C totals against defect-code totals. |
| `Worksheet_SelectionChange` | Publish | Locks the Publish sheet — clicking outside `C4:C7` redirects the cursor to A1, preventing accidental config edits. Set `DeveloperMode = True` in the code to bypass. |

---

## 🔧 Helper Functions

| Function | Description |
|----------|-------------|
| `EvalMath(rng)` | Extracts and sums all numbers from a range using Regex. Returns `-` if total is zero. Used for defect-code totals. |
| `EvalPower(rng)` | Same as `EvalMath` but returns `0` instead of `-`. Used for A/B/C and power-cut totals where numeric zero is needed for comparisons. |
| `CleanString(txt)` | Trims spaces and non-breaking spaces before Model/Part lookups, preventing missed matches due to formatting differences. |
| `IgnoreNumberAsTextErrors()` | Suppresses Excel's "Number stored as text" and "Inconsistent formula" warnings across the active sheet. |

---

## ⚙️ Configuration (`Publish` sheet)

| Cell | Setting | Example |
|------|---------|---------|
| `C4` | Report date | `15.04.2026` |
| `C5` | PowerPoint output path | `C:\Users\...\15.04.2026.pptx` |
| `C6` | Summary slide number | `7` |
| `C7` | Path to `COST.xlsx` | `C:\Morning meeting report template\COST.xlsx` |

---

## 📝 Notes

- **Rejection categories:** A = Major · B = Minor · C = Rework *(or as defined by your process)*
- **Defect codes 1–10** map to defect types defined in your quality system
- `COST.xlsx` must be accessible at the configured path before using cost sync buttons
- Macros must be enabled on open — the workbook will not function without them
- The PowerPoint template must already exist at the output path before publishing

---

## 📁 External Dependencies

```
COST.xlsx                    ← Unit cost lookup (Model + Part Name → Cost)
Morning meeting PPT template ← Pre-formatted .pptx the macro writes into
```

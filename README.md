# VBA-loop-formatting-headers


# Excel VBA Automation: Header Formatting Loop

This repository contains a simple Excel VBA automation tool that applies consistent header formatting across all sheets in a workbook. The project demonstrates how to use a VBA loop to automate repetitive formatting tasks.

## 📂 Files in This Repository

- `CleanUpData.xlsm` – Macro-enabled Excel file containing the VBA script
- `RawData.xlsx` – Sample raw data used for testing
- `Sample.png` – Example output showing formatted headers

## ⚙️ Features

- Automates header formatting for **all worksheets** in an Excel workbook
- Combines **macro recording** and **manual VBA scripting**
- Easy to use and modify for similar automation tasks

## 🔁 How It Works

1. First, the macro recorder is used to capture the steps for:
   - Adding headers
   - Formatting headers

2. The recorded process is wrapped into a loop using Visual Basic for Applications (VBA), so the formatting can be applied to every worksheet.

### 🧠 VBA Code Sample

```vb
Public Sub CleanUpData()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        Worksheets(i).Select
        AddHeaders
        FormatHeaders
        i = i + 1
    Loop
End Sub
````

> Note: `AddHeaders` and `FormatHeaders` should be defined as separate subroutines in your VBA module.

## 🚀 How to Use

1. Open the `CleanUpData.xlsm` file in Excel.
2. Navigate to the **Developer** tab.
3. Click on **Visual Basic** → open `Module 1` to view the code.
4. Or go to **Developer** → **Macros** → run the `CleanUpData` macro to apply the formatting.
5. See `Sample.png` for a visual example of the expected output.

## 📧 Contact

* **LinkedIn**: [Max Nguyen Hoang Minh](https://www.linkedin.com/in/max-nguyen-hoang-minh)
* **Email**: [maxnguyenhoangminh@gmail.com](mailto:maxnguyenhoangminh@gmail.com)

---

Feel free to fork or contribute if you find this useful!

```

---

Would you like help creating the `AddHeaders` and `FormatHeaders` subroutines too?
```

---
title: How to Fix Excel 2021 Formulas Not Working Properly | Step-by-Step Guide | Stellar
date: 2024-05-19T18:32:11.585Z
updated: 2024-05-20T18:32:11.585Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2021 Formulas Not Working Properly | Step-by-Step Guide
excerpt: This article describes How to Fix Excel 2021 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair corrupt .xlb,repair .xltm,repair .xls files,repair corrupt .xlsm files,repair damaged .xltx,repair .xltx,repair damaged .xltm,repair corrupt excel,repair corrupt .csv files
thumbnail: https://www.lifewire.com/thmb/he26hNDyU2vpo9vDwma05qKk-R8=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/espnplus-6d79c347f2df42d7a38cdb0ba9a2472b.jpg
---

## How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide

**Summary:** Excel formulas sometimes fail to function correctly and even return an error. This article explains what you might be doing wrong that prevents Excel formulas from working properly and solutions to resolve the issue. If your formulas have disappeared from the Excel spreadsheet and you are having trouble recovering them, you can use an Excel repair tool to recover the formulas.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When working with Excel formulas, situations may arise when the formula doesn’t calculate or update automatically. Or, you may receive errors by clicking on a formula.

## Problems Causing the ‘Excel Formulas not Working Properly’ Issue and Solutions

Let’s check out the possible reasons that cause Excel formulas to work properly and solutions to resolve the issue.

### Problem 1 – Switching Automatic to Manual Calculation Mode

Automatic and manual are the two modes of calculation in Microsoft Excel.

By default, Excel is set to automatic calculation mode. Everything is recalculated automatically when any changes are made in a worksheet in this mode. You may switch from automatic to manual mode to disable the recalculation of formulas, particularly when working with a large Excel file with too many formulas.

Excel will not calculate automatically when set to manual calculation mode. And this may make you think that the Excel formula is not working properly.

### Solution – Change Calculation Mode from Manual to Automatic

To do so, perform these steps:

- Click on the column with problematic formulas.
- Go to the **Formulas** tab, click the **Calculation Options** drop-down, and select **Automatic**.

![Automatic to Manual Calculation Mode](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/switch-manual-to-automatic-calculation-1.png)

### Problem 2 – Missing or Mismatched Parentheses

It’s easy to miss or incorrectly place parentheses or include extra parentheses in a complex formula. If a parenthesis is missing or mismatched and you click Enter after entering a formula, Excel displays a message window suggesting to fix the issue (refer to the screenshot below).

![Missing or Mismatched Parentheses](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/missing-or-mismatched-parantheses-2.png)

Clicking ‘Yes’ might help fix the issue. But Excel might not fix the parentheses properly, as it tends to add the missing parentheses at the end of a formula which won’t always be the case.

### Solution – Check for Visual Cues When Typing or Editing a Formula with Parentheses

When typing a formula or editing one, Excel provides visual cues to determine if there’s an issue with the parentheses inserted in a formula. Checking for these visual cues can help you fix missing/mismatched parentheses.

- Excel helps identify parenthesis pairs by highlighting them in different colors. For instance, the pair of parenthesis outside is black.
- Excel does not make the opening parentheses bold. So, if you’ve inserted the last closing parentheses in a formula, you can determine if your parentheses are mismatched.
- Excel helps identify parentheses pairs by highlighting and formatting them with the same color once you cross over them.

### Problem 3 – Formatting Cells in an Excel Formula

When adding a number in an Excel formula, don’t add any decimal separator or special characters like $ or €. You may use a comma to separate a function’s argument in an Excel formula or use a currency sign like $ or € as part of cell references. Formatting the numbers may prevent the formula from functioning correctly.

### Solution – Use Format Cells Option for Formatting

Use Format Cells instead of using a comma or currency signs for formatting a number in the formula. For instance, rather than entering a value of $10,000 in your formula, insert 10000, and click the ‘Ctrl+1’ keys together to open the Format Cells dialog box.

![Format Cells Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-3-1.png)

### Problem 4 – Formatting Numbers as Text

Numbers are displayed as left-aligned in a sheet in a worksheet, and text formatted numbers are right-aligned in cells. Excel considers numbers formatted as text to be text strings. Thus, it leaves those numbers out of calculations. As a result, a formula won’t work as intended. For example, in the following screenshot, you can see that the SUM formula works correctly for normal numbers. But, when the SUM formula is applied to numbers formatted as text, the formula doesn’t return the correct value.

![Cells Formatted as Text](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/cells-formatted-as-text-4.png)

Sometimes, you may also see an apostrophe in the cells or green triangles in the top-left corner of all the cells when numbers in those cells are formatted as Text.

### Solution – Do Not Format Numbers as Text

To fix the issue, do the following:

- Select the cells with numbers stored as text, right-click on them, and click Format Cells.
- From the Format Cells window, click on Number and then press OK.

![Format Cells in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-5.png)

### Problem 5 – Double Quotes to Enclose Numbers

Avoid enclosing numbers in a formula in double-quotes, as the numbers are interpreted as a string value.

Meaning if you enter a formula like =IF(A1>B1, “1”), Excel will consider the output one as a string and not a number. So, you won’t be able to use 1’s in calculations.

### Solution – Don’t Enclose Numbers in Double Quotes

Remove any double quotes around a number in your formula unless you want that number to be treated as text. For example, you can write the formula mentioned above as “1” =IF(A1>B1, 1).

### Problem 6 – Extra Space at Beginning of the Formula

When entering a formula, you may end up adding an extra space before the equal (=) sign. You may also add an apostrophe (‘) in the formula at times. As a result, the calculation won’t be performed and may return an error. This usually happens when you use a formula copied from the web.

### Solution – Remove Extra Space from the Formula

The fix to this issue is pretty simple. You need to look for extra space before the equal sign and remove it. Also, ensure there is an additional apostrophe added in the formula.

## Other Things to Consider to Fix the ‘Excel Formulas not Working Properly’ Issue

- If your Excel formula is not showing the result as intended, see this [blog](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).
- When you refer to other worksheets with spaces or any non-alphabetical character in their names, enclose the names in ‘single quotation marks’. For example, an external 5reference to cell A2 in a sheet named Data enclose the name in single quotes: **‘Data’!A1**.
- You may see the formula instead of the result if you have accidentally clicked the ‘Show Formulas’ option. So, click on the problematic cell, click on the Formula tab, and then click Show Formulas.
- If you’re getting an error “Excel found a problem with one or more formula references in this worksheet”, find solutions to fix the error [here](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

## Conclusion

This blog discussed some problems you might make causing an Excel formula to stop working properly. Read about these common problems and solutions to fix them. If a problem doesn’t apply in your case, move to the next one. If you cannot retrieve formulas in your Excel sheet, using an [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) like Stellar Repair for Excel can help you restore all the formulas. It does so by repairing the Excel file (XLS/XLSX) and recovering all the components, including formulas.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)




## Filter Not Working Error in Excel [Fix 2024]

**Summary:** The filter is not working issue in Excel can occur due to several reasons, like blank rows, hidden rows, merged cells, corrupted data, etc. In this post, we will mention the reasons why the filter is not working correctly in Excel and several fixes to resolve the issue. We will also mention an advanced Excel repair tool to repair the Excel file if corruption in file is the cause of the issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

You can use the Filter function in Excel to filter data in large-sized Excel files quickly. While using Excel filters, sometimes, you face a situation where the filter is disabled or may fail to function properly.

![Filter Option Disabled](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/filter-option-disabled-1024x112.jpg)

The Excel filter usually fails to work if you have not selected the complete and correct range of data. Let’s learn more about the “Sort and Filter not working in Excel” issue and look at the possible methods to fix it.

## **Why the Filter is not Working in Excel?**

You can face the “filter is not working” issue if you are applying the filter on a protected worksheet or trying to find the data from a hidden row. Besides this, there could be many other reasons contributing to this issue, such as:

- The data you are trying to filter is in merged cells.
- The Excel file automatically selected the data up to the first empty cell, excluding the remaining rows.
- Grouped sheets in Excel file.
- Blank row in the Excel sheet.
- You are trying to apply a filter on an invalid data range.
- The workbooks in which you’re facing the filter issues are corrupted.
- You are specifying incorrect criteria in the filter columns.

## **Solutions to Resolve the Filter is not Working Issue in Excel**

There might be two scenarios: the Excel filter option is disabled/grayed out or the filters fail to function properly. You can follow the given troubleshooting solutions to resolve the issue based on the scenario you’re facing.

## **Scenario 1 – Filter Option is Disabled or Grayed Out**

### **Method 1: Check and Un-group the Worksheet**

When you apply filters to a single sheet in a grouped set, Excel disables the filter option in other sheets within the group. You can check the grouped sheets and try ungrouping them to enable the filter option. Here’s how to do so:

- In the Excel file, go to the **Group** section.

![Excel file navigation: Accessing the Group section
](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-group-section-1024x114.jpg)

- Right-click on the **Ungroup Sheets.**

Alternatively, you can press the Shift + Alt + Left keys to ungroup the sheets.

### **Method 2: Unprotect Worksheet**

The “disabled Excel filter” issue can also occur if your worksheet is protected. You can unprotect the worksheet to enable the filter option. To do so, go to the **Review** tab and then select **Unprotect Sheet.**

![Excel file: Navigating to Group section, resolving 'disabled Excel filter' issue with worksheet protection, unprotecting sheet from Review tab for filter activation.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-review-and-select-unprotect-sheet-1024x115.jpg)

### **Method 3: Check and Uninstall Excel Add-ins**

Sometimes, the Excel filter gets disabled due to faulty or corrupted Excel add-ins. You can run the Excel in Safe mode to check whether the issue has occurred due to add-ins. To do this, type excel /safe in the Run window and click **OK.**

![Troubleshooting disabled Excel filter caused by add-ins: Running Excel in Safe mode with 'excel /safe' in Run window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/type-excel-safe-command.jpg)

In safe mode, if you see the filter option, it indicates some problematic Excel add-ins were causing the issue. In such a case, you can check and uninstall the faulty Excel add-ins to fix the issue.

## **Scenario 2 – Filter is not Working**

### **Method 1: Try Clearing Filters**

Sometimes, the Excel filter fails to work correctly if some filters from the previous sessions are still active. In such a case, you can clear the applied filters. Follow the below steps:

- In Excel file, click Sort & Filter option.
- Select clear.

![Excel: Clicking 'Sort & Filter' and selecting 'Clear' option.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-clear-option.jpg)

### **Method 2: Select Entire Data**

The filter not working issue in Excel can occur when the range selected for filtering is incomplete or incorrect. You need to make sure that you’ve selected the entire data range in Excel. You can use the Ctrl+A keys to select the entire content in the worksheet.

### **Method 3: Check and Delete Blank Cells from the Table’s Columns**

When you apply a filter to the data, Excel expects data to be in a continuous range. Excel filters do not consider the blank cells, thereby resulting in incorrect functioning of the filter. To resolve this issue, check and delete all blank cells. In case your Excel file is too large to delete the blank cells, then you can add a “Serial number” row as an alternative. Adding serial number row creates a data continuity, thus helping in fixing the filter-related issue.

### **Method 4: Unhide Hidden Rows and Columns**

Hidden rows or columns in worksheets can also affect the filter functionality. You can check and unhide rows/columns to troubleshoot the issue. Here is how to do so:

- In the affected Excel file, go to Home.
- Click on **Format > Hide & Unhide**.

![Excel file: Navigating to Home, accessing Format > Hide & Unhide.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-format-select-hide-or-unhide-option-1024x228.jpg)

- Click **Unhide Rows** or **Unhide Columns** (as required).

![Selective unhiding in Excel: 'Unhide Rows' or 'Unhide Columns' as needed.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-unhide-rows-unhide-columns.jpg)

### **Method 5: Unmerge Cells**

You can experience the filter in Excel is not working issue if you are using the filter to extract data from merged cells. Ensure to unmerge the “merged cells” before applying a filter in Excel. Follow the below steps to unmerge the merged cells in Excel:

- Navigate to the **Home** option.
- In the toolbar, select the **Merge & Center** option.
- Click **Unmerge Cells.**

### **Method 6: Repair the Workbook**

Sometimes, the **Filter Not Working in Excel** issue can occur due to inconsistencies in file structure. If these issues occurred due to corruption in the worksheet, you can repair it using the Open and Repair tool. It is an in-built tool in Excel that is used to repair corrupted Excel files. Here are the steps to use this tool:

- In the Excel application, navigate to the **File** option.
- Click **Open** and then click **Browse** to choose the Excel file.
- In the **Open** dialog box, click the problematic Excel file.
- Click the arrow next to the **Open** option and select **Open and Repair.**
- Click **Repair** to recover as much data as possible.
- The application prompts a message after the repair process is complete. Click **Close**.

In most cases, the Open and Repair tool can easily fix corruption issues in the Excel file. However, for any reason, if the [open and repair tool doesn’t work](https://www.stellarinfo.com/blog/ms-excel-open-and-repair-option-is-not-working/) you can consider repairing the file using a professional Excel Repair tool. Stellar Repair for Excel is one such advanced and secure tool to repair Excel files. With this tool’s powerful scanning capabilities, you can repair highly corrupted Excel files and recover all their objects with complete integrity. The tool is compatible with all Windows editions, including the latest Windows 11.

## **Closure**

Several reasons are associated with the **filter not working issue in Excel**. The filter option may not work as expected if you have not selected the complete and correct range of data or for many other reasons. You can follow the troubleshooting methods discussed above to fix the issue. If the filter fails to work due to corruption in the workbook, then try [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can even repair severely damaged files. It also helps to recover all the data from corrupted files without changing the original formatting. You can check the tool’s functionality by downloading its demo version. It allows you to preview all the repairable objects in the corrupted Excel file.


## [Fix] Excel formula not showing result

**Summary:** Is your Excel spreadsheet showing text of a formula you’ve entered and not its result? This blog explains the possible reasons behind such an issue. Also, it describes solutions to fix the ‘Excel formula not showing result’ error. You can try Stellar Repair for Excel software to recover engineering and shared formulas.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Sometimes, when you type a formula in a cell of worksheet and press Enter, instead of showing the calculated result, it returns the formula as text. For instance, Excel cell shows:

![Excel not Showing Formula](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-not-showing-formula-1.jpg)

But you should get the result as:

![Excel Formula Working Sample](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-not-showing-formula-2.jpg)

## **Why Does Excel Show or Display the Formula Not the Result?**

**Following are the possible reasons that may lead to the ‘Excel showing formula not result’ issue:**

1. You accidentally enabled “**Show Formulas**” in Excel.
2. The cell format in a spreadsheet is set to text.
3. ‘Automatic calculation’ feature in Excel is set to manual.
4. Excel thinks your formula is text (Syntax are not followed).
5. You type numbers in a cell with unnecessary formatting.

## **How to Fix ‘Excel Showing Formula Not Result’ Issue?**

### **Solution 1 – Disable Show Formulas**

If only the formula shows in Excel not result, check if you have accidentally or intentionally enabled ‘show formula’ feature of Excel. Instead of applying calculations and then showing results, this feature displays the actual text written by you.

You can use the ‘Show Formulas’ feature to quickly view all formulas, but if you are not aware of this feature, and enabled it accidentally, it can be a headache. To disable this mode, go to ‘**Formulas**’ and click on ‘**Show formula enabled**.’ If it’s previously enabled, it will be disabled by just clicking on it.

![Show Formula Enabled](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/show-formulas-1.jpg)

### **Solution 2 – Cell Format Set to Text**

Another possible reason that only formula shows in Excel not result could be that the cell format is set to text. This means that anything written in any format in that cell will be treated as regular text. If so, change the format to General or any other. To get Excel to recognize the change in the format, you may need to enter cell edit mode by clicking into the formula bar or just press F2.

![Enter Cell Edit Mode by Clicking into the Formula Bar](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/formula-f2.jpg)

### **Solution 3 – Change Calculation Options from ‘Manual’ to ‘Automatic’**

There is an “automatic calculation” feature in Excel, which tells Excel to do calculations automatically or manually. If ‘Excel formula is not showing results’, it may be because the automatic calculations feature is set to manual. This issue is not easily detected because it results in calculating formula in one cell but if you copy it to some other cell, it will retain the first calculation and will not recalculate on the base of the new location. To fix this, follow these steps:

- In Excel, click on the ‘**File**’ tab on the top left corner of the screen.
- In the window that opens, click on ‘**Options**’ from the left menu bar.
- From ‘Excel Options’ dialog box, select ‘**Formulas**’ from the left side menu and then change the ‘**Calculation options**’ to ‘**Automatic**’ if it’s currently set as ‘Manual’.

![Automatic Calculations Feature](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/calculation-options.jpg)

- Click on ‘**OK**’. This will redirect you to your sheet.

### **Solution 4 – Type Formula in the Right Format**

There is a proper way to tell Excel that your text is a formula. If you don’t write the formula in a particular format, Excel considers it as simple text and hence no calculations are performed according to it. For this reason, keep the following in mind when typing a formula:

- **Equal sign:** Every formula in Excel should start with an equal sign (=). If you miss it, Excel will mistake your formula as regular text.

- **Space before equal sign:** You are not supposed to enter any space before equal sign. Maybe a single space will be hard for us to detect, but it breaks the rule of writing formulas for Excel.

- **Formula wrapped in quotes:** You need to make sure that your formula is not wrapped in quotes. People usually make this mistake of writing a formula in quotes, but in Excel, quotes are used to signify text. So your formula won’t be evaluated. But you can add quotes inside formula if required, for example: =SUMIFS(F5:F9,G5:G9,”>30″).

- **Match all parentheses in a formula:** Arguments of Excel functions are entered in parenthesis. In complex cases, you may need to enter more sets of parenthesis. If those parentheses are not paired/closed properly, Excel may not be able to evaluate the entered formula.

- **Nesting limit:** If you are nesting two or more Excel functions into each other, for example using nested IF loop, remember the following rules:
  - Excel 2019, 2016, 2013, 2010, and 2007 versions only allow to use up to 64 nested functions.
  - Excel 2003 and lower versions only allow up to 7 nested functions.

### **Solution 5 – Enter Numbers without any Formatting**

When you use a number in the formula, make sure you don’t enter any decimal separator or currency sign, e.g. $, etc. In an Excel formula, a comma is used to separate arguments of a function and a dollar sign makes an absolute cell reference. Most of these special characters have built-in functions so avoid using them unnecessarily.

## **What to Do If the Manual Solutions Don’t Work?**

If you’ve tried out the manual solutions mentioned above but still unable to resolve the ‘Excel formula not showing result’ issue, you can try repairing your Excel file with the help of an automated [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as **Stellar Repair for Excel**.

This reliable and competent software scans and repairs Excel files (.XLSX and .XLS). It also helps recover all the file components, like formulas, cell formatting, etc. Armed with an interactive GUI, this software is extremely easy to work with, and its advanced algorithms allow it to fend off Excel errors with ease.

[![Free Download for windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Conclusion**

This blog outlined the possible reasons that may cause ‘Excel not showing formula results’ issue. Check out these reasons and implement the manual fixes, depending on what resulted in the problem in the first place. If none of these fixes help resolve the issue, corruption in the Excel file might be preventing the formulas from showing the actual results. In that case, using **Stellar Repair for Excel** tool might help.


## How to fix Microsoft Excel not responding error and save your data

**Summary:** This guide helps you resolve Excel not responding and frequent Excel freeze issues in Excel on Windows 10. It mentions some effective solutions to repair Excel and resolve Excel is not responding problem. These solutions will also help you fix Excel crashing problem while working on the spreadsheet.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Similar to any other program, you may experience problems with Microsoft Excel while opening or working on a document. Sometimes, it may not start at all or freeze and display an error message such as ‘Excel is not responding’. When it happens, you may want to wait for the program to respond.

![Microsoft Excel is not Responding ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/Microsoft-Excel-not-responding.png)

‘Microsoft Excel is not responding’ problem

**Tip:** If you are experiencing Excel not responding problem with a particular Excel file, it’s quite possible that the file is corrupt or partially damaged. And thus, leading to an Excel freeze or crash problem. Use Stellar Repair for Excel software to quickly repair and restore Excel (.xls/.xlsx) file in its original, intact form. You can download the free trial version of the software from the below link.

But if Excel doesn’t respond after a while and remains stuck, you need to force close the program from “**Task Manager**”. Now, this could be disastrous if happens while you are working on an important Excel document that took you hours to prepare. Force closing Excel due to such error can damage the Excel document and it may fail to open next time.

## **Why Excel is Not Responding?**

Excel may stop responding, freeze, or crash suddenly due to several reasons. It can happen while saving a spreadsheet or opening an Excel document. It may also occur while editing or inserting images, graphs, etc. But usually, it occurs when the system crashes or shuts down abruptly while you are working on a document. Here’s an instance,

_Suppose, you worked overnight on a critical document which is to be presented at a meeting the next day. This Excel spreadsheet includes critical graphs and charts, and much more. When you are about to save it, there is a power failure, and your system shuts down without warning. When the power is up, you restarted the system to check your Excel. To your dismay, a message pops up – “Excel Crashed” or “Microsoft Excel not responding”._

This could be frustrating. However, there is no need to despair as there are solutions to not just overcome this error but other corresponding issues such as Excel freezing, hanging, crashing, etc. Below is an infographic that quickly briefs all the possible solutions to fix Excel not responding error.

## **Solutions to Fix ‘Microsoft Excel is not responding’ Error**

Follow the solutions discussed below in the given order to fix Excel freezing and hanging issues.

### **Solution 1**: **Open Excel in Safe Mode**

If Excel is not working as intended and frequently stops responding, you may try to start Excel in Safe Mode. It is a common DIY way to fix ‘Excel is not responding’ problem.

In Safe Mode, Excel starts with only essential services, bypasses certain functionalities and doesn’t load the add-ins, which might be the reason behind the error in MS Excel . To open and troubleshoot Excel in **Safe Mode**:

- Press **Windows + R** keys, type **excel.exe /safe and press ‘Enter’ or click ‘OK’**

![MS Excel in Safe Mode](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/02/Excel-safe-mode-1024x523.png)

MS Excel in Safe Mode

Open the Excel file and check if it still crashes. If not, the problem could be a faulty add-in or formatting and styling error.

Proceed to the next solution to check and fix the problem.

### **Solution 2: Check for Faulty and Unwanted Add-ins**

**In Microsoft Excel, there are two types of add-ins:**

- **COM add-ins**
- **Other Add-ins Installed as XLAM, XLA, or XLL File**

**Both types of add-ins can cause the freezing problem in Excel . Follow the steps below to disable unwanted and faulty add-ins:**

- In Excel , click **File and g**o to **Options** to open ‘**Excel Options**’ window
- Click **Add-ins** button to view and manage ‘**Microsoft Office Add-ins**’
- Uncheck required add-ins to disable them
- At this stage, you can also click the ‘**Remove**’ button to remove any unwanted add-ins

![Disable COM Add-Ins](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/02/Remove-COM-Addins.jpg)

Disable COM Add-Ins

- Now enable an add-in and check the Excel performance. Observe Excel for not responding error or freezing problem

If Excel doesn’t freeze, enable subsequent add-in and then again use Excel to observe it. Repeat the steps until you find the faulty plugin, which is causing the problem.

Then remove it from Excel add-ins to resolve the problem.

### **Solution 3****:** **Install the latest Windows and Office Updates**

This problem may also occur if Windows and MS Office are not updated. Therefore, install the latest updates for both Microsoft Windows and Microsoft Office.

You can set the installation and update option to ‘**Automatic mode**’ in Windows. This will download and install critical updates for MS Office, which might fix the Excel performance issue. The steps to enable automatic updates are as follows:

1. Go to **Settings> Update & Security> Windows Update**

![Enable Automatic Windows updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/update-windows-1024x796.jpg)

Enable Automatic Windows updates

- Click **Advanced options** and enable all the toggle switches to automatically download and install updates for Windows and other Microsoft products

![Update Microsoft Products](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/upload-WIndows-Products-1024x796.jpg)

Update Microsoft Products

After update, restart Excel and check if the problem is resolved.

**_NOTE:_** _From now on, MS Excel will also get the latest update consistently, without the need for manual intervention._

### **Solution 4:  Check and Disable Anti-virus**

Antivirus is important for device safety. However, if your antivirus conflicts with MS Office apps such as Excel, it could lead to Excel freezing and not responding errors.

To check if the problem is due to anti-virus, disable it and reopen the Excel document. Check if Excel performs well or if it still hangs.

![Example of Antivirus conflict with Microsoft Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/Antivirus.jpg)

Example of Antivirus conflict with Microsoft Excel

If the problem is resolved, contact your antivirus software provider for help to keep antivirus running without affecting the system and other programs such as MS Excel.

### **Solution 5: Change the ‘Default Printer’**

Although it may seem irrelevant, changing the default printer is another easy and effective solution to overcome the error. Reason being, Excel communicates with the printer to find supported margins when we open an Excel sheet.

If Excel doesn’t find the supported margin, it may stop responding or crash. The steps to change the default printer are as follows:

1. Open **Control Panel** on your Windows system
2. Click **Printer and Devices**
3. Right-click **Microsoft XPS Document Writer** to set it to the default printer

![Change in Default Printer Setting](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/02/Default-Printer-Setting-1.jpg)

Change in Default Printer Setting

Reopen the Excel document to check whether the error occurs or not.

### **Solution 6: Repair Microsoft Office**

A corrupt or damaged Microsoft Office can also cause the ‘Excel is not responding’ problem. You can resolve this by repairing the Microsoft Office files. The steps are as follows:

1. **Close** all running MS Office programs
2. Go to **Control Panel** on your Windows system
3. Click **Programs** and then **Programs** **and** **Features**
4. Select **Microsoft Office** and in the Microsoft Office window, click ‘**Change**’
5. Then select the ‘**Repair**’ option and click ‘**Continue**’

![Repair MS Office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/02/Repair-MS-Office.jpg)

Repair MS Office

This may take a while. After the repair is done, check your Excel program and file for the error.

### **Solution 7: Remove and Reinstall Microsoft Office**

Sometimes, repairing MS Office may not work. In such a case, removing and reinstalling Microsoft Office can resolve the ‘Excel is not responding’ problem. To do so, follow these steps:

1. **Close** all running MS Office programs
2. Go to **Control Panel** on your Windows system
3. Click **Programs** and then **Programs** **and** **Features**
4. Right-click on **Microsoft Office** and choose **Uninstall**

![ Uninstall MS Office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/02/Uninstall-MS-Office.jpg)

Uninstall MS Office

Then run the MS Office installation setup to re-install MS Office on your system.

### **Solution 8: Repair Microsoft Excel (XLS/XLSX) file**

In several situations, a corrupt or partially damaged Excel (XLS/XLSX) file is the cause of this error. In such a case, you can download and install **Stellar Repair for Excel** to repair the corrupt or damaged Excel file. By repairing the Excel file, you can resolve the Excel freezing error quickly without applying much efforts.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**The steps to use the software for Excel file repair are as follows:**

1. Download, install and launch the [**Excel file repair software**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)
2. Browse and select the corrupt Excel file

![Stellar Excel repair software](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/Stellar-Excel-repair-software-1024x576.jpg)

- Click ‘**Repair’** to start repairing the damaged Excel file
- After file repair, it provides a preview. Check your file

![Excel file repaired](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/Excel-repair-software-to-fix-Excel-errors-1024x576.jpg)

- Then click the ‘**Save File**’ option in the main menu
- You can either choose default location or browse a new folder location to save the repaired Excel file

![Save repaired Excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/Save-repaired-excel-file-1024x576.jpg)

- After repair, open the file in Excel and continue with your work

![Repaired Excel file saved](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/04/Repaired-Excel-file-is-sved-1024x576.jpg)

And keep Stellar Repair for Excel installed on your system. You never know when you might need this handy tool.

You may also refer to Microsoft support for more details on [Excel not responding, hangs, freezes or stops working](https://support.office.com/en-us/article/excel-not-responding-hangs-freezes-or-stops-working-37e7d3c9-9e84-40bf-a805-4ca6853a1ff4) issues.

## **Conclusion**

Now that the methods for fixing the ‘Excel is not responding’ error are before you, try all these and see which one works for you. If the cause of this error is a damaged or corrupt Excel file, only repairing the XLS/XLSX file can resolve the issue.

For this purpose, it’s recommended to use a reliable software such as _Stellar Repair for Excel_ as it offers an easy-to-use interface, thereby making Excel file repair process a seamless experience.

The software recovers table, chart, chart sheet, cell comment, number, text, shared formulas, image, formula, sort and filter, and other objects. It also preserves worksheet properties, layout, and cell formatting. It can repair multiple XLS/XLSX files simultaneously and fix all [Excel file corruption](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) errors.

All these features extend the software capabilities beyond just fixing the ‘Excel not responding’ error.


## How to fix Microsoft Excel has stopped working error?

**Summary:** This blog discusses the possible reasons behind ‘Microsoft Excel has stopped working’ error and solutions to resolve the error manually. You can use Stellar Repair for Excel to quickly repair the file and recover all its data in a hassle-free manner.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Has your Microsoft Excel program stopped working or is acting strange? Excel not responding is a common issue you may experience on launching the application or opening a spreadsheet.

![Microsoft Excel has stopped working](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/07/Excel-has-stopped-working.jpg)

Figure 1 – Microsoft Excel Has Stopped Working Error Message

## **Possible Causes behind ‘Microsoft Excel has Stopped Working’ Error, and Solutions Thereof**

_**Note:** Several users have reported about encountering the ‘_**_Excel has stopped working’ issue on Windows 10, 8, and 7 OS_** _after installing an update for Excel (KB3118373). If you too have installed the update, then uninstall it and check if it solves the error. For detailed information, refer to this_ [link](https://docs.microsoft.com/en-us/office/troubleshoot/excel/excel-has-stopped-working-error)_._

## Get Rid of corrupt Excel File

**Summary:** What to do when an Excel file is corrupted? This is a common question that is often asked by Microsoft Excel users. If you too are seeking an answer to this question, read the blog to learn about a few manual workarounds and a specialized Excel file repair tool to resolve the Excel file corruption issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

An Excel file gets corrupted due to various reasons such as a virus/malware attack, sudden system shutdown when the Excel file is still open, power failure while working with an Excel spreadsheet, etc.

When Microsoft Excel detects corruption in a workbook, it attempts to repair the workbook by starting _‘File Recovery mode’._

**Tip!** If the file recovery mode doesn’t start, you may use the manual repair process or an Excel repair tool, such as Stellar Repair for Excel to repair a corrupted Excel file. The software can help you quickly retrieve contents from a damaged, corrupt, or inaccessible Excel file and restore the file to its original state.

There even exist a few manual tips that can be used to recover data from damaged MS Office Excel files.

## **Workarounds to Use When an Excel File is Corrupted**

**_Note:_** _Before carrying out any of the repair and recovery workarounds, it is advised that you must save a backup copy of the damaged file. This is to prevent your files from turning completely inaccessible in case the methods fail to give desired results._

### **Workaround 1: Use the Open and Repair Method**

If MS Excel cannot repair a corrupted workbook automatically, you can try to do it manually. To do so, perform the following:

- Open the corrupt file, like you normally open any file, by clicking **File** > **Open**.
- Browse and locate the folder containing the corrupted document.
- When the Open dialog box is displayed:
- Select the Excel document.
- Click on the arrow present to the right side of the Open button and select **Open and Repair** option.

![open and repair excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/open-and-repair.jpg)

Figure 1 – Open and Repair Feature

If this doesn’t help repair the broken Excel file or you encounter  [Open and Repair does not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) issue, proceed with the next workaround.

**_Tip!_** _Try an alternative solution, i.e._ **_Stellar Repair for Excel software_** _to repair and recover corrupt Excel files (.xlsx or .xls) when the ‘Open and Repair’ method won’t work._  

### **Workaround 2: Restore an Excel File with a Shadow Copy**

If you’re a **Windows 7 or Vista user,** you can try restoring the corrupted spreadsheet by using a shadow copy (or a previous version). [Shadow copy](<https://en.wikipedia.org/wiki/Shadow_Copy#:~:text=Shadow%20Copy%20(also%20known%20as,the%20Volume%20Shadow%20Copy%20service>.) is basically a snapshot (backup copy) of computer files or volumes. The snapshot may contain an older version of your Excel file that has become damaged now. To find out, do the following:

- **Launch File Explorer**, and right-click the folder in which the **file is saved.**
- Choose **Properties.**
- Look for and click the **Previous Versions** tab. This will display a list of entries under **Folder versions** or **File versions**, going back a few days or weeks.
- Double-click one with a date when the file was accessible and could be read. Then, try to open its older version. If it opens, save the older version with a new name and execute the procedure with new file/folder entries.

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Shadow-copy.png)

Figure 2 – Volume Shadow Copy

You would have to repeat the process until you reach the point where the file became damaged. With this, you will get a baseline version of the file, but data may still have been lost.

### **Workaround 3: Test your Assumptions**

If you receive a message saying “[**Excel file corrupted and cannot be opened**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)”, you would probably believe it. However, there could be other reasons besides corruption that may cause Excel to throw this error message.

Your Office suite, which Excel is a part of, maybe having some primary issues in it causing problems while opening one Excel document. So, try opening another Excel file to check if the problem exists with all the files or just one.

If other Excel documents work correctly, it means that only the particular document is corrupt. On the contrary, if the issue is with your Office suite, repairing the current Office installation may help fix the issue. For this, perform these steps:

- Go to **Control Panel and click Uninstall** **the Program**.
- Choose **Office.**
- Click Change, and hit the **Repair button.**

You can **reinstall** the entire Office package. Once reinstalled, try to open the file to check if the issue has been fixed and the **Excel file repaired.**

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Repair-MS-Office.jpg)

Figure 3 – MS Office Repair

### **Workaround 4: Use Excel File Repair Tool**

If the above manual solutions fail, use [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to successfully repair your damaged Excel workbook and recover all its data. Essentially, the software rebuilds damaged Excel workbook data at a granular level to recover every single object & all the original properties of the workbook.

**Suggested Read:** [**How to repair corrupt Excel files using Stellar Repair for Excel?**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Why Use Stellar Repair for Excel Software?**

- Repairs severely corrupted XLSX and XLS files.
- Can handle corrupt Excel files of any size.
- Demo version allows previewing recoverable Excel file items for free.
- **Supports Microsoft Excel 2019** and all lower versions.
- Compatible with **Windows 10 and lower versions.**
- Tested and recommended by **Microsoft Excel MVPs.**

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Final Word**

When an Excel file is corrupted, it won’t open at all or you won’t be able to access all the file data. Such a situation can lead to unnecessary halts, impacting work productivity.

There are manual workarounds that may help fix the corrupt Excel file and recover its data, such as the ones covered in this blog. However, these solutions might not work in severe corruption cases and may require technical assistance. Also, they may result in some data loss.

To overcome the limitations of manual workarounds, it is recommended to go for a professional Excel file repair tool such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It helps **repair corrupt Excel** (XLS or XLSX) files and restores all worksheet data, such as the table, chart, chart sheet, cell comment, sort and filter, image, formula, etc. in a few simple clicks. Moreover, the software provides a free preview of the recoverable data with its demo version. You can check the preview to evaluate how the software works.


## How Do I Restore Excel File ?

Excel (XLS/XLSX) file corruption is a common issue that may arise due to problems, such as abrupt system shutdown, system crash, malware or virus intrusion, unsupported add-in, etc. However, MS Excel comes with a recovery feature that allows users to recover or restore Excel files lost or corrupted when the MS Excel quits unexpectedly or computer crashes due to hardware or software-related issues.  

But sometimes, the restored version of the Excel file may not contain all the data or changes you made recently. In such cases, you are left with the option to repair the damaged or corrupt Excel file with the help of an Excel repair software, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

[![](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Solutions to Restore Excel File

Below we have discussed various solutions to repair, recover, and restore unsaved, lost, corrupt, or inaccessible Excel files with the help of MS Excel options and Excel repair software.

**_NOTE:_** _Some of these methods work only if you had enabled or configured certain required settings or options before corruption or losing the Excel file._

### 1\. Restore Lost or Unsaved Excel File via ‘Recover Unsaved Workbooks’ Option

The _Recover Unsaved Workbooks_ option allows you to restore the Excel file that you forgot to save before closing MS Excel or it stopped working. The steps are as follows:

- Navigate to ‘_File > Open._’
- Scroll to the bottom and click on ‘Recover Unsaved Workbooks.’ This opens a File Explorer browser window.

![recover unsaved excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/recover-unsaved-document.png)

- Now look for the Excel file that you lost and then double-click on it to open and restore the Excel file.
- Once the file is open, click on **File > Save As** and save it to your desired location.

However, this option works only if you had enabled and set up the ‘_Save AutoRecover information every X minutes_’ option and enabled the ‘_Keep the last AutoRecovered version if I close without saving’_ option.  

![change excel settings](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/change-settings-excel.png)

Also, this may not restore the most recent copy of your Excel file, but it will surely save your time that went into creating and editing the workbook.

### 2\. Repair and Restore Corrupt or Inaccessible Excel File

If the Excel file you are trying to open or access is corrupt, MS Excel will display an error message. When MS Excel detects corruption, it starts ‘File Recovery Mode’ to repair the Excel workbook.

However, if it doesn’t start automatically, you can manually repair minor corruption error and restore the Excel file by following these steps:

- Navigate to **File > Open** and then navigate to the corrupt or damaged Excel file location.
- Select the file and then click on the arrow beside the ‘**Open**’ button to choose the **‘Open and Repair…**’ option.

![open and repair excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/open-and-repair-excel-file.png)

- Now click ‘**Repair’** if you want to recover maximum data from a corrupt or inaccessible Excel workbook and restore it to a new Excel file.

![repair corrupt excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/repair-corrupt-excel-sheet.png)

If the ‘**Repair’** option fails, use the Extract Data option that extracts values, data, and formulas from the Excel workbook.

### 3\. Revert Excel File to Last Saved Version

When an Excel file gets corrupt or damaged while working, you may revert the Excel file to the last saved version. This allows you to restore the Excel file version that isn’t corrupt or damaged. The steps are as follows,

- Click File > Open.
- Double-click the name of the workbook that you have opened in Excel.
- Click Yes to reopen the workbook.

The workbook opens without any changes you’ve made that might have caused the workbook to become corrupted.

### 4\. Use Excel Repair Software to Restore Excel File

If the corrupt or damaged Excel file isn’t repaired by using MS Excel options or you lost data after restoring the unsaved version or after reverting to the last version, you can rely on an Excel repair software, such as Stellar Repair for Excel.

The software repairs Excel files with 100% integrity and restores the Excel data, such as table, Pivot tables, charts, etc., while preserving the workbook properties and formatting. The software is safe as it doesn’t alter or make changes to the original Excel file and saves the repaired data to a new Excel workbook at your desired location.

It supports Excel files (XLS and XLSX), created in MS Excel 2019, 2016, or older versions. It resolves Excel corruption errors, such as Unable to open Excel file, in a few clicks.

To repair a corrupt or inaccessible Excel workbook by using Stellar Repair for Excel software, follow these 3 simple steps,

- Select the corrupt or damaged Excel workbook file (XLS/XLSX).

![stellar repair for excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/stellar-repair-for-Excel-choose-xls-xlsx-file.png)

- Repair and then preview the Excel workbook.

![repaired excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/repaired-excel-file.png)

- Save the workbook at default or your desired location.

![save repaired excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/save-repaired-excel-file.png)

Once saved, you can open and start editing or working on the Excel sheet. Make sure to enable AutoRecovery and AutoSave options to avoid data loss due to corruption.

## Conclusion

In this article, we discussed methods and solutions to restore Excel file when it gets damaged or corrupt due to various issues. Although Excel can help you recover or restore Excel file after corruption, it can resolve only minor corruption errors. Thus, it may fail if the Excel file corruption is severe. In such cases, an Excel repair software, such as Stellar Repair for Excel, comes in handy. It resolves Excel corruption, repairs tables, charts, chart sheets, Pivot table, etc., and restores them to a new Excel sheet with 100% consistency.


<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="8358498916"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>
<ins class="adsbygoogle"
    style="display:block"
    data-ad-format="autorelaxed"
    data-ad-client="ca-pub-7571918770474297"
    data-ad-slot="1223367746"></ins>

<span class="atpl-alsoreadstyle">Also read:</span>
<div><ul>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-infinix-hot-40-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Infinix Hot 40 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-mini-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 12 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-realme-narzo-60-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Realme Narzo 60 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-itel-a60-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Itel A60</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-p40plus-by-fonelab-android-recover-video/"><u>How to restore wiped videos on P40+</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-realme-c55-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Realme C55 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-samsung-galaxy-a05s-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Samsung Galaxy A05s to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-rog-phone-7-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from ROG Phone 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-sony-xperia-5-v-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Sony Xperia 5 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-ios-beta-version-on-iphone-12-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS Beta Version on iPhone 12? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-manually-install-a-driver-in-windows-11-by-drivereasy-guide/"><u>How to Manually Install a Driver in Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-camon-20-pro-5g-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Camon 20 Pro 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-7-plus-without-data-loss-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 7 Plus without Data Loss? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-s18-pro-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on S18 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-s17-pro-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to S17 Pro Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-iphone-xs-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Data from iPhone XS? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-scalp-and-copy-trading-on-mt4-and-mt5-by-mt4copier-guide/"><u>How to scalp and Copy Trading on MT4 and MT5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-honor-90-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Honor 90</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-vivo-x90s-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Vivo X90S?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-honor-magic-6-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Honor Magic 6 Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-realme-10t-5g-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Realme 10T 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-mov-files-on-xiaomi-by-aiseesoft-video-converter-play-mov-on-android/"><u>How to play MOV files on Xiaomi ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-12-pro-max-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 12 Pro Max | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-samsung-galaxy-a24-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Samsung Galaxy A24 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-avchd-mts-files-on-razr-40-by-aiseesoft-video-converter-play-mts-on-android/"><u>How to play AVCHD MTS files on Razr 40?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-vivo-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Vivo</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-samsung-galaxy-a25-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Samsung Galaxy A25 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-11-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 11 Pro | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-13-without-losing-any-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 13 without Losing Any Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-xiaomi-13t-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Xiaomi 13T using Video Repair Utility?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-solve-mkv-lagging-problem-in-xperia-1-v-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How to solve MKV lagging problem in Xperia 1 V?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-xr-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone XR | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-tecno-pova-5-pro-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Tecno Pova 5 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-redmi-k70-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Redmi K70</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-drivers-with-windows-device-manager-in-windows-11-and-10-and-7-by-drivereasy-guide/"><u>How to identify some outdated drivers with Windows Device Manager in Windows 11 & 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-7-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 7 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-vivo-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Vivo using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-honor-x7b-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Honor X7b</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-tecno-spark-10-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Tecno Spark 10 5G? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-moto-g84-5g-by-fonelab-android-recover-video/"><u>How to recover old videos from your Moto G84 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-oneplus-ace-3-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from OnePlus Ace 3 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-contacts-from-oppo-a78-by-fonelab-android-recover-contacts/"><u>How to get back lost contacts from Oppo A78.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-infinix-note-30-vip-racing-edition-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Infinix Note 30 VIP Racing Edition Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6s-to-an-older-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6s to an Older iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-insert-sign-in-xltm-files-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to insert sign in .xltm files</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-honor-90-lite-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Honor 90 Lite to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-an-mp4-on-xiaomi-redmi-note-13-proplus-5g-by-aiseesoft-video-converter-play-mp4-on-android/"><u>How to play an MP4 on Xiaomi Redmi Note 13 Pro+ 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-13c-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi 13C</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-rog-phone-7-ultimate-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted ROG Phone 7 Ultimate Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-edge-40-pro-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Edge 40 Pro.</u></a></li>
<li><a href="https://techidaily.com/how-to-reset-apple-iphone-7-plus-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset Apple iPhone 7 Plus to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-in-2024-unbiased-reviews-of-youtube-mp3-converters-for-beginners/"><u>New In 2024, Unbiased Reviews of YouTube MP3 Converters for Beginners</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-can-i-get-more-stardust-in-pokemon-go-on-nokia-c210-drfone-by-drfone-virtual-android/"><u>How can I get more stardust in pokemon go On Nokia C210? | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/updated-how-to-resize-flv-videos-efficientlywindows-mac-android-iphone-and-online-for-2024/"><u>Updated How to Resize FLV Videos EfficientlyWindows, Mac, Android, iPhone & Online for 2024</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/can-t-view-mov-movies-content-on-motorola-by-aiseesoft-video-converter-play-mov-on-android/"><u>Can’t view MOV movies content on Motorola </u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-network-locked-sim-card-inserted-on-your-oneplus-ace-2-phone-unlock-it-now-by-drfone-android/"><u>In 2024, Network Locked SIM Card Inserted On Your OnePlus Ace 2 Phone? Unlock It Now</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/guide-to-mirror-your-xiaomi-redmi-note-12t-pro-to-other-android-devices-drfone-by-drfone-android/"><u>Guide to Mirror Your Xiaomi Redmi Note 12T Pro to Other Android devices | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-can-i-do-animated-logo-design-without-design-skills-for-2024/"><u>Updated Can I Do Animated Logo Design Without Design Skills for 2024</u></a></li>
<li><a href="https://location-social.techidaily.com/does-oneplus-open-have-find-my-friends-drfone-by-drfone-virtual-android/"><u>Does OnePlus Open Have Find My Friends? | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-dose-life360-notify-me-when-someone-checks-my-location-on-honor-magic-6-drfone-by-drfone-virtual-android/"><u>In 2024, Dose Life360 Notify Me When Someone Checks My Location On Honor Magic 6? | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-fake-the-location-to-get-around-the-mlb-blackouts-on-asus-rog-phone-7-ultimate-drfone-by-drfone-virtual-android/"><u>In 2024, Fake the Location to Get Around the MLB Blackouts on Asus ROG Phone 7 Ultimate | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/how-to-reset-your-samsung-galaxy-s24-ultra-lock-screen-password-by-drfone-android/"><u>How to Reset your Samsung Galaxy S24 Ultra Lock Screen Password</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-in-2024-top-10-sketch-animation-video-makers/"><u>Updated In 2024, Top 10 Sketch Animation Video Makers</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-share-location-in-messenger-on-infinix-note-30-5g-drfone-by-drfone-virtual-android/"><u>How to Share Location in Messenger On Infinix Note 30 5G? | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-9-mind-blowing-tricks-to-hatch-eggs-in-pokemon-go-without-walking-on-google-pixel-fold-drfone-by-drfone-virtual-android/"><u>In 2024, 9 Mind-Blowing Tricks to Hatch Eggs in Pokemon Go Without Walking On Google Pixel Fold | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-resolve-your-apple-iphone-se-2020-keeps-asking-for-outlook-password-by-drfone-ios/"><u>In 2024, Resolve Your Apple iPhone SE (2020) Keeps Asking for Outlook Password</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-how-can-we-unlock-our-samsung-galaxy-s23-ultra-phone-screen-by-drfone-android/"><u>In 2024, How Can We Unlock Our Samsung Galaxy S23 Ultra Phone Screen?</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-pgsharp-save-you-from-ban-while-spoofing-pokemon-go-on-tecno-camon-20-pro-5g-drfone-by-drfone-virtual-android/"><u>How PGSharp Save You from Ban While Spoofing Pokemon Go On Tecno Camon 20 Pro 5G? | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-top-4-sim-location-trackers-to-easily-find-your-lost-samsung-galaxy-a05s-device-by-drfone-android/"><u>In 2024, Top 4 SIM Location Trackers To Easily Find Your Lost Samsung Galaxy A05s Device</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/learn-how-to-lock-stolen-your-apple-iphone-x-properly-by-drfone-ios/"><u>Learn How To Lock Stolen Your Apple iPhone X Properly</u></a></li>
<li><a href="https://unlock-android.techidaily.com/forgotten-the-voicemail-password-of-vivo-v30-lite-5g-try-these-fixes-by-drfone-android/"><u>Forgotten The Voicemail Password Of Vivo V30 Lite 5G? Try These Fixes</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/new-the-ultimate-list-10-best-webcam-recorders-for-windows-10/"><u>New The Ultimate List 10 Best Webcam Recorders for Windows 10</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/5-ways-to-reset-tecno-spark-10c-without-volume-buttons-drfone-by-drfone-reset-android-reset-android/"><u>5 Ways to Reset Tecno Spark 10C Without Volume Buttons | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-text-messages-from-vivo-y100a-to-new-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Text Messages from Vivo Y100A to New Phone | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/ultimate-guide-from-apple-iphone-15-icloud-activation-lock-bypass-by-drfone-ios/"><u>Ultimate Guide from Apple iPhone 15 iCloud Activation Lock Bypass</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-i-transferred-messages-from-lava-blaze-2-5g-to-iphone-12xs-max-in-seconds-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How I Transferred Messages from Lava Blaze 2 5G to iPhone 12/XS (Max) in Seconds | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-migrate-android-data-from-infinix-hot-30i-to-new-android-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Migrate Android Data From Infinix Hot 30i to New Android Phone? | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/what-is-fake-gps-location-pro-and-is-it-good-on-nokia-xr21-drfone-by-drfone-virtual-android/"><u>What is Fake GPS Location Pro and Is It Good On Nokia XR21? | Dr.fone</u></a></li>
</ul></div>



---
title: How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide | Stellar
date: 2024-05-19T18:32:11.472Z
updated: 2024-05-20T18:32:11.472Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide
excerpt: This article describes How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair excel 2019,repair damaged .csv,repair excel 2016,repair excel 2021,repair damaged .xls,repair .xltx files,repair damaged .xlsm files,repair .xlsm,repair corrupt .csv files,repair damaged .xlsm,repair .csv files
thumbnail: https://www.lifewire.com/thmb/RHk5CzUskZEHtVQS5Kba30nHhvY=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/is-kindle-unlimited-worth-it-fda01dceb923406a8524c64d2b72693e.jpg
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




## \[Solved\] : How to Fix MS Excel Crash Issue

Microsoft [Excel may stop responding](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), hang, freeze, or stop working due to several reasons, such as in compatible add-ins. In such a case, you may receive one of the following error messages.

- Excel has stopped working

![Excel has stopped working](https://www.stellarinfo.com/public/image/catalog//article/email-repair/exchange/excelnew1.jpg)

- Excel is not responding

![Excel is not responding](https://www.stellarinfo.com/image/catalog/article/excelnew2.jpg)

- A problem caused the program to stop working correctly. Windows will close the program and notify you if a solution is available.

![A problem caused the MS Excel to stop working correctly](https://www.stellarinfo.com/image/catalog/article/excelnew3.jpg)

## Why Does Excel Keep Crashing?

If Excel keeps crashing on your PC while opening a workbook, saving Excel file, scrolling or editing cells, etc., it indicates a problem with your Excel program or the Excel file.

Microsoft Excel may crash due to any one or more reasons given below,

-  Incompatible Add-Ins
- Outdated MS Excel program
- Conflict with other programs or antivirus tool
-  Excel file created by third party software
- Problem with network connection
-  Combination of Cell formatting and stylings
- Problem with MS Office installation
- Partially damaged or corrupt Excel file

## Problems Caused by Excel Crash Issue

Microsoft Excel crash may cause damage to Excel file and also lead to Excel (XLS/XLSX) file corruption.

Such corrupt Excel files can't be opened or accessed via MS Excel app. If you try to access a corrupt Excel file, MS Excel may fail to open the file or stop responding and crash. Additionally, you may receive the following or similar error message,

![Excel files can't be opened or accessed](https://www.stellarinfo.com/image/catalog/article/excelnew4.jpg)

In such a case, you should immediately try to recover the Excel file. You may do so by restoring the Excel file from backup or by using an [Excel File Repair software.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) Otherwise, continue following this guide.

## How to Solve Excel Crash Issue?

Before heading to solutions, follow these troubleshooting steps to resolve the Excel Crash issue.

### **Step 1: Copy File to Local Drive**

If you are trying to access and edit or view an Excel file from a network drive, try moving the file to local drive. This will help you find if there is something wrong with the file or the network.

### **Step 2: Ensure Sufficient Memory**

Excel files can grow fairly large when you start adding lots of formatting and shapes. Make sure that your system has enough RAM to run the application.

![Ensure Sufficient Memory](https://www.stellarinfo.com/image/catalog/article/excelnew5.jpg)

If you often work with large Excel files and complex data values& formulas, then install 64-bit versions of MS Office. It will give you an advantage of larger processing capacities and prevent Excel from crash or freeze.

### **Step 3: Check If Excel is Open and In Use by Another Process**

Open **Task Manager** and close all processes or apps (tasks) that may be using or have access to your Excel file that you are working on. You can find this detail in status bar of Excel program at the bottom of program window.

![Task Manager](https://www.stellarinfo.com/image/catalog/article/excel6.jpg)

After closing the tasks, try to access the Excel file and check if this fixes the performance and crash problem in Excel.

### Step 4: Test and Repair Excel File

Create a copy of the Excel file and install **Stellar Repair for Excel** software. It's free to download. Scan and repair your Excel file using the software. After repair, save the Excel file at your desired location and then open the Excel file in the MS Excel program.

![Stellar Repair for Excel software](https://www.stellarinfo.com/image/catalog/article/excel7.jpg)

This should ideally fix all the issues with Excel.

However, if the Excel program still crashes, the problem lies within the system or program. Follow the solutions discussed in this guide to try to fix the Excel crash issue.

**NOTE:** To save repaired Excel file using the mentioned software, you must purchase the activation key and activate it.

## Solutions to Fix MS Excel Crash Issue

Following are some solutions to resolve problems with MS Excel such as,

- Excel not responding
- Excel won't open
- Excel keeps crashing

Follow these solutions in the given order. In case a method doesn't work, move to the next one.

### Solution 1: Restart Excel in Safe Mode

By starting MS Excel in safe mode, you can run the program without loading the Excel add-ins and with limited features. But COM add-ins are excluded.

To launch Excel in safe mode, close MS Excel and follow these steps,

- Create a shortcut of MS Excel (.exe) on Desktop
- Press and hold the Ctrl key while launching the program
- Click 'Yes' when a prompt appears to confirm

Alternatively, press Windows+R, type excel /safe and press 'Enter'. Use this to open Excel in safe mode on Windows 10, 8.1, 8, or 7 system.

![type excel /safe](https://www.stellarinfo.com/image/catalog/article/excel8.jpg)

Now try to open and access the Excel file and check if the issue is resolved. If it's not, head on to the next solution.

### Solution 2: Check and Remove Faulty Add-ins

In case Excel doesn't crash in Safe Mode, it's possible that some faulty add-ins are the culprit behind frequent Excel crash and freeze. These Excel add-ins may interfere or conflict with the Excel program.

![Check and Remove Faulty Add-ins](https://www.stellarinfo.com/image/catalog/article/excel9.jpg)

Find and remove the faulty add-in. It can resolve the issue. To do so, follow these steps,

- Restart Excel in normal mode and go to File> Options> Add-ins
- Choose COM Add-ins from the drop-down and click Go

![COM Add-ins](https://www.stellarinfo.com/image/catalog/article/excel10.jpg)

- Uncheck all the checkboxes and click OK

![Uncheck all the check boxes](https://www.stellarinfo.com/image/catalog/article/excel11.jpg)

- Restart Excel and check if the issue is resolved
-  If Excel doesn't crash or freeze anymore, open COM Add-ins and enable one add-in at a time followed by Excel restart. Then observe Excel for freeze or crash problem

This will help you find out the faulty add-in, which is causing the problem. Remove the add-in which is causing the problem to resolve the issue. If that doesn't fix, move to the next solution.

### Solution 3: Check and Install the Latest Updates

If you haven't set Windows to Download and Install Updates automatically, do it now.

Apart from updating the operating system, latest Windows updates sometimes fixes bugs for other applications installed on the system such as MS Office. Often installing an important update that you might have missed may correct the Excel crash problem.

You can also update MS Office manually. Follow these steps,

Go to File > Account

 Under Product Information, select Update Options and click Update Now

![Product Information](https://www.stellarinfo.com/image/catalog/article/excel12.jpg)

If you have installed MS Excel from Microsoft Store, open the store and update your Office applications.

NOTE: This also works if you can't open Excel file or Excel crashes after Windows upgrade from Windows 7 or Windows 8/8.1 to Windows 10.

After installing the latest MS Office updates, check if Excel works fine. If not, head to the next solution.

### Solution 4: Clear Conditional Formatting Rules

If a sheet is causing Excel to freeze or crash, there might be a problem with that particular sheet. In such a case, you may try clearing the Conditional Formatting rules. The steps are as follows,

- Under Home, click 'Conditional Formatting > Clear Rules\> Clear Rules from Entire Sheet'

![Conditional Formatting](https://www.stellarinfo.com/image/catalog/article/excel13.jpg)

- You may repeat this step for all other sheets in the Excel workbook
- Then click File> Save as and save the Sheet as a new file at a different location

This avoids overwriting or making changes to the original Excel file. Once done, try working on the sheet.

If this doesn't work out, move to the next solution.

### Solution 5: Remove Multiple Cell Formatting and Styles

If a workbook is being shared and edited by others on different platforms then it's possible that many cells are formatted differently. This can cause issues with Excel such as crash and freeze. It can also lead to Excel file corruption. The problem mostly occurs when a workbook contains multiple worksheets using different formatting.

You can [follow this guide](https://docs.microsoft.com/en-gb/office/troubleshoot/excel/too-many-different-cell-formats-in-excel) to remove different cell formats and styles, and then open the Excel file.

### Solution 6: Disable Microsoft Excel Animation

Animations require additional processing power and resources. By disabling animations in Excel, you may resolve Excel freeze and crash issue. This also improves MS Excel performance.

To disable the animations in MS Excel, follow these steps:

- Go to File > Options
- Click 'Advanced' and check 'Disable hardware graphics acceleration'animation

![Disable hardware graphics acceleration](https://www.stellarinfo.com/image/catalog/article/excel14.jpg)

- Click 'OK' to close the window and then restart MS Excel

This has helped many users in fixing the Excel crash issue. If it doesn't work for you, head to the next solution.

### Solution 7: Check If Excel File is Generated by a Third-Party Application

There are applications which you may have used to generate Excel files to fetch data. For instance, downloading data from Google Analytics in Excel format.

Sometimes, these Excel files are not generated correctly by such third-party apps. Thus, some features in Excel may not work as intended when you access the files in MS Excel.

In such a case, you should get in touch with the app developer for help with the file or use Stellar Repair for Excel to repair such Excel files.

### Solution 8: Check If Antivirus or Other Apps are Conflicting with MS Excel

Ensure your antivirus is up-to-date and not conflicting with MS Excel. An outdated antivirus tool may conflict with Excel which can cause the application to hang, freeze, or crash.

- Update your antivirus
- Try disabling the add-in or integration between Excel and antivirus. See if it works

Alternatively, you may disable the anti-virus tool temporarily to check if it is the culprit behind Excel performance issue and crash. If that resolves the problem, get in touch with your antivirus vendor and report the problem.

They might provide you with a better solution or workaround to fix this problem without disabling the antivirus protection.

IMPORTANT NOTE: Disabling or altering antivirus protection makes your PC vulnerable to malicious attacks and virus or malware intrusion.

### Solution 9: Clean Boot Windows to Inspect the Cause Behind Excel Crash

When Windows boot, it starts several processes, services, and application during start up automatically, which runs in the background.

These startup apps and services can interfere with other applications such as MS Excel. To find out if that's the cause behind Excel crash, you can perform a Clean Boot.

This helps you identify processes, services, or applications that are conflicting with Excel. Steps to perform Clean Boot are as follows,

- Press Windows key + R, type MSConfig, and press 'Enter'
- In System Configuration window, click on the General tab and choose Selective startup

![System Configuration](https://www.stellarinfo.com/image/catalog/article/excel15.jpg)

Uncheck 'Load startup items' and click 'OK'

After this, close all running applications and restart your PC

Check if the crash problem with Excel is resolved. Uninstall the conflicting apps or update them. If your issue is not resolved, follow the next solution.

### Solution 10: Repair or Reinstall MS Office

Repairing Office programs may also resolve Excel crash issues if caused by damaged MS Excel program or MS Office files. The steps are as follow,

- Close all MS Office apps and open the Control Panel
- Click Uninstall a program under Programs

![Uninstall a program](https://www.stellarinfo.com/image/catalog/article/excel16.jpg)

- Click on Microsoft Office and then click on the Change option
- Choose 'Quick repair' and then select 'Repair'
- Click 'Continue' to repair MS Office installation

You may also try 'Online Repair' if this fails to fix the issue. After repair, if the Excel issue persists, reinstall MS Office.

## Need More Help?

If none of the above-mentioned solutions worked for you, it indicates that the problem is not with the Excel program but with the Excel file. If you haven't tried the Stellar Repair for Excel software, do it now.

Select the Excel file which is causing the problem and repair it with the software. It's a powerful Excel repair software that can fix all the problems with Excel files (XLS/XLSX). It repairs corrupt and severely damaged Excel files.

The software is compatible with all Excel files created using MS Excel 2019, 2016, 2013, 2010, 2007, 2003 or 2000.

After repairing and saving the Excel file, you can open it in your MS Excel program and work on it without any performance issue. To know more about this software, visit [this page.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)


## Fix Cannot Paste the Data Error in Microsoft Excel

When copying and pasting the cell contents or attributes from one workbook to another, you can encounter the "Cannot paste the data" error. The error indicates that the values you are trying to paste do not match the cell format (Currency, Date, Text, etc.) used in the column. The error can occur if you are attempting to paste information into a merged cell. You can also get this error if the data you are trying to paste contains more columns than the destination worksheet can accept or if the copy area and paste area aren't the same size.

## Why you cannot Paste the Data in Microsoft Excel?

You can experience the "Excel cannot paste the data" error due to one of the following reasons:

- Locked cells
- Data size limitation
- Formatting issues
- The file size is too large
- Overloaded clipboard
- Data type mismatch
- Invalid range
- Lack of permissions
- Excel file is corrupted
- Compatibility issues
- While pasting the merged cells

## **Methods to Fix MS Excel Cannot Paste the Data Error**

The "Microsoft Excel cannot paste the data" error usually appears when copying and pasting data within the spreadsheet. It primarily affects the copy-and-paste feature within the Excel file. However, the impact of this error may vary with the data you are working with. When this error occurs, first try restarting the Excel application. Ensure that you've saved all your important data before restarting the application to prevent data loss. If, after performing this basic step, you still face the issue, then try the following troubleshooting methods.

### **Method 1: Check and Unlock Cells in Excel**

You can get the Excel cannot paste the data error if the cells you are trying to paste in the Excel file are locked. You can check and unlock the specific cells in the Excel file using the below steps:

- Go to the **Home** tab in your Excel file and click **Format Cell Font** popup launcher.

![Clicking Cell Font In Home](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/cannotpastedata/click-format-cell-font-on-home-tab.jpg)

- In the **Format Cells** dialog box, click **Protection**.

![Selecting Protection from Font family with formatting](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/cannotpastedata/click-protection-in-format-cells.jpg)

- Under the **Protection** option, unselect the locked field and then click **OK**.

![Click on Locked Cell under Protection Tab](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/cannotpastedata/select-locked-under-protection.jpg)

### **Method 2: Check and Fix the Cell Format**

Sometimes, you can get the "Cannot paste the data" error in Excel if there is an issue with the formatting of the cells in the column (in which you are trying to paste the data). You can check whether the cell formats of the columns of the source file (from where you are copying data) match the destination file (in which you are pasting the cell data). If not, then change the cell format. Here's how to do so:

- In the Excel file, click on the affected column heading (whose cells you need to modify).
- Navigate to the **Home** tab and click the **General** dropdown menu.

**![Navigate to Home and Click general drop down menu in Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/cannotpastedata/click-general-dropdown-menu.jpg)**

- Click on the desired cell format that matches the values you are trying to paste into the column.

### **Method 3: Check Copy and Paste Area Size**

The "data cannot be pasted" issue can also occur if Excel detects the copy area of the range of cells is not of the same size and shape as that of the paste area. So, make sure the size of the range of cells for the paste area is same as the copied area. To ensure this, you can try the following:

- **Use Upper-left Cell**

 You can select the upper-left cell instead of the complete range to paste.

- **Use Paste Special Option**

- You can use the **Paste Special** option to paste the data. To use this option, go to **Home > Paste Special**.

![Navigate to Home And click on Paste Special in Excel](https://www.stellarinfo.com/public/image/catalog/article/Repair-Office-Documents/cannotpastedata/go-to-home-and-then-click-paste-special.png)

- You can choose from the different options, such as paste all, formulas, operations, Transpose, skip blanks, etc. to take control over how the data is pasted.

![Click on Paste and Selection of Operations from Paste Special Properties](https://www.stellarinfo.com/public/image/catalog/article/Repair-Office-Documents/cannotpastedata/select-operations-from-paste-special.jpg)

### **Method 4:** Check and Unmerge Cells

The merged cells in the workbook can create discrepancies when copying and pasting data. These can create mismatch between the source and destination cell areas, resulting in errors. Make sure you are not pasting the data copied from merged cells. You can check and unmerge the cells using the below steps:

- Open the Excel file (in which you are copying data) and go to the Home tab.
- Click Merge & Center > Unmerge Cells.

![Under Excel Navigate to Home and click on Merge and Center](https://www.stellarinfo.com/public/image/catalog/article/Repair-Office-Documents/cannotpastedata/go-to-home-click-on-merge-and-center.jpg)

### **Method 5: Unselect the Excel DDE (Dynamic Data Exchange) Option**

Sometimes, you can get the Excel cannot paste the data error while copying and pasting data from Excel file to external data sources or applications. It usually appears if you have established DDE links between them. To resolve this, you can uncheck the "Ignore other applications that use Dynamic Data Exchange (DDE)" option. Follow the below steps to do so:

- Open your Excel file.
- Go to **File > Options**.

In **Excel Options**, click **Advanced** and unselect **"Ignore other applications that use Dynamic Data Exchange (DDE)**. Click **OK.**

**![Under Excel Navigate to Home and click on Merge and Center](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/cannotpastedata/click-ignoring-dynamic-data-exchange-option.jpg)**

## Here are some additional solutions you can try to fix the "Cannot paste the data" error in Excel

### **Check the Excel File Format**

Sometimes, the copy and paste error can occur if you try to move data from an older version of Excel file (XLS) into a newer version file (XLSX) or vice versa. You can try to change or convert the file format to resolve the issue.

### **Clear Clipboard**

Too many items on the clipboard can prevent you from copying and pasting the data. You can [empty your clipboard](https://support.microsoft.com/en-au/office/clear-the-clipboard-7afbf55f-d7d5-4096-87a0-eb17f821d321) by deleting all clips. To do this, go to **Home** and click the **Clipboard** option.

### **Change the Column's Cell Format**

You can get the "Cannot paste the data" error when the information you need to paste does not match the cell format in the column. In such a case, you can change the column's cell format. To do this, navigate to the **Home** option and then click on the **General menu** arrow. Then, change the cell format suitable to the type of information you are trying to paste in that column.

### **Add Columns**

The copy-and-paste error in Excel can also occur if there are not enough columns to accept the data that you are trying to paste. You can insert more columns and then try copying and pasting data. To insert additional columns, click on the column heading, click **Home**, and then select **Insert.**

### **Remove Incompatible Add-ins**

Sometimes, incompatible or faulty add-ins can also create conflicts with Excel's clipboard operations, leading to the error. To check if add-ins are causing the issue, open the Excel file in safe mode. In safe mode, if you are able to paste the data, then check and uninstall the incompatible Excel Add-ins.

## **What to do if Nothing Works?**

If the above methods do not work, then corruption in the Excel file could be the cause of the "Cannot paste the data" error. In such a case, you can use Excel's built-in Open and Repair utility to repair the corrupted file. Here's how to use this utility to repair the Excel file:

- Open your Excel application.
- Navigate to **File** and then click **Open.**
- Click the **Browse** option to select the affected Excel file.
- The **Open** dialog box opens up. Click on the corrupted file.
- From the **Open** dropdown, select the **Open and Repair** option.
- The Excel will prompt you to choose one of the below options:
- Repair
- Extract
- Cancel
- Select the **Repair** option to extract maximum data from the file. If the Repair option fails, choose the **Extract** option to recover the data, excluding formulas and values.

 The [Open and Repair utility may not be able to repair your Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) if it is severely corrupted. In such a situation, the best possible way out is to use a professional Excel repair software to repair the file. You can try Stellar Repair for Excel. It can repair highly damaged or corrupted Excel (.xls, .xlsx, .xltm, .xltx, or .xlsm) file and recover all the objects from the file by maintaining the original formatting. The tool supports 2007 and higher editions of MS Excel. You can download the demo version of the software to preview the repaired objects and verify its functionality.

## Conclusion

The "Excel cannot paste the data" error in Excel can occur due to different factors. You can check and clear the clipboard, unmerge the cells, unlock the cells, and use various other methods mentioned above to troubleshoot the error. If you fail to copy and paste the data due to file corruption, then opt for an advanced Excel repair software, like [Stellar Repair for Excel.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) It is a reliable tool that can help you repair highly damaged/corrupted Excel files. It enables you to recover all the objects from the corrupted Excel file, with complete integrity.


## Excel File Corruption Warnings and Solutions

**Summary:** Many users reported error messages they receive when they try to save or open an Excel file. In this blog, you will learn about the warning messages that indicate your Excel file is corrupt and possible solutions to repair it. It also outlines the Stellar Repair for Excel to repair corrupt Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Excel users often report about receiving warning messages suggesting corruption in the workbook. This usually happens while opening an Excel file, ‘.xls’ or ‘.xlsx’ file created by earlier versions, or attempting to create a copy of the workbook.

Excel file corruption may occur due to several reasons including (but not limited to) virus infection, sudden system shutdown during write operation, and leaving excel file open on the shared network.

### **Occurrences of Excel File Corruption Warnings**

_Occurrence 1 – “Excel found unreadable content in <filename>. Do you want to recover the contents of this workbook? If you trust the source of this workbook, click Yes”._

![Image of Excel Found Unreadable Content error message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Excel-found-unreadable-content-error.png)

On clicking ‘Yes’, you will receive the following error:

 _“The file is corrupt and cannot be opened”._

![Image Of Excel File Corruption error Message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/The-file-is-corrupt-and-cannot-be-opened-error.png)

_Occurrence 2 – “Excel cannot open the file <filename>, because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file”._

![Image of Excel File Format Or Extension is Not Valid error message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Excel-file-extension-error.png)

Besides the warning messages outlined above, there are a few other tell-tale signs of Excel file corruption such as:

- Excel crashes or freezes, preventing you from accessing the workbook and information stored in it.
- Unexpected errors occur during the save operation listed as below:
  - _“An unexpected error has occurred. AutoRecover has been disabled for this session of Excel”._
  - _“Errors were detected while saving <filename>”._

### **Solutions to Fix Excel File Corruption Issue**

Follow the below-listed solutions to deal with corruption issues in Excel:

**NOTE:** If you encountered problem opening Excel files after upgrading to latest Windows Operating System (OS) and Office program, try updating your Office as well as Windows OS to latest patches provided on the Microsoft site. Microsoft frequently releases Office and Windows OS patches to help users’ correct known errors. Check if you can open the corrupt workbook after installing the update.

#### **Solution 1 – Use Open and Repair Utility**

Excel comes with a built-in recovery mechanism. It automatically starts ‘File Recovery Mode’ when a user opens a corrupt workbook, and attempts to open and repair the workbook. Sometimes, the recovery mode might not start automatically. In that case, you will need to repair the Excel file manually by using ‘[Open and Repair](https://support.office.com/en-us/article/repairing-a-corrupted-workbook-7abfc44d-e9bf-4896-8899-bd10ef4d61ab)’ utility.

**Steps to use Microsoft’s built-in repair utility are as follows:**

**Step 1:** Select **File** > **Open**.

**Step 2:** Click the folder containing the corrupt workbook, and then click **Browse**.

**Step 3:** In the **Open** window, select the corrupt workbook.

**Step 4:** Next, click the arrow in the **Open** button, and then click **Open and Repair**.

![Image of Open and Repair in-built utility](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Open-and-Repair.png)

**Step 5:** In the window that appears, click **Repair**.

![Image of Excel warning message after using open and repair in-built utility.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Open-and-Repair-Repair-option.png)

If  [‘Open and Repair’ doesn’t work in excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), select **Extract Data** to extract formulas and values from the corrupt workbook.

**_NOTE:_** _If you need a quick solution to salvage your data, use an Excel file repair tool._

Or else, attempt the following solutions to deal with [corruption in Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

#### **Solution 2 – Uninstall and Re-install Office Installation**

**_NOTE:_** _Make sure to create a backup of your Excel file before uninstalling and re-installing your Office application._

Download the Office uninstall support tool to remove the application.

You can read: [Simple Ways to Open Corrupt Excel file Without any Backup](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**To reinstall Microsoft Office, follow these steps:**

_**NOTE:** Before proceeding with Office re-installation process, make sure that you have license keys ready._

**Step 1:** Open the [Microsoft Office](http://www.office.com/) site.

**Step 2:** Select **Sign in**.

**_NOTE:_** _You may skip this step if you’re already signed in._

**Step 3:** After signing in, from the Office sign-in page, click **Install**/**Install Office**

Your Office application will get re-installed. Now open the backed-up Excel file and see if the problem is fixed.

#### **Solution 3 – Move Excel File to a Different Location**

Often moving a corrupt Excel file to a different location can help solve the corruption problem. Here’s how:

**Step 1:** Open the corrupt Excel file by navigating to the following path:

**C:\\Users\\User\_Name\\AppData\\Roaming\\Microsoft\\Excel**

_**NOTE**: Make sure to replace User\_Name with your user name. If you are unable to find the Excel file, you will have to search for the file manually in Program Files (x86)._

![Image of Moving Excel File to a Different Location](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/program-files.png)

**Step 2:**  Open the Excel folder, and move the corrupt file to some other location.

**Step 3:** Delete the files from the Excel folder.

Now try opening the Excel file you have moved and see if the issue is resolved.

#### **Solution 4 – Use Excel File Repair Software**

If none of the above solutions works for you, use **Stellar Repair for Excel**. It is a specialized [Excel file repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that helps repair corrupt Excel file and recover workbook data in its original state.

Essentially, the software helps rebuild the corrupt file to restore every single object in the file. It can recover objects including user-defined charts, conditional formatting rules, formatting of the charts, properties of worksheet, engineering formulas, etc.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

#### **Steps to use Stellar Repair for Excel are as follows:**

**Step 1:** Download, install and launch **Stellar Repair for Excel** software.

**Step 2:** In **Select File** window, click **Browse** to select the file you want to repair.

![Image of Stellar Excel Repair software start screen.
Click on Select File -> Browse](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/user-interface-1024x544.png)

_**NOTE:** If you are unaware of the Excel file location, click ‘Search’ in the Select File window to find the file._

**Step 3:** Once the files are selected, click **Repair** to initiate the repair process.

![Image of Repair Process window after selecting the files to be repaired](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/select-file.png)

**Step 4:** Preview the repaired file and select all or specific files you want to save.

![Image of Preview of Repaired File ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Preview-1024x545.png)

**Step 5:** Click **Save File** on **Home** menu.

![Image of Save File Button on Home Menu.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Save-file.png)

**Step 6:** In **Save File** window, choose ‘Default Location’ or ‘Select New Folder’ to select the location where you wish to save the file. Click **OK**.

![Image of save File window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Save-file.jpg)

The selected files will be saved at the specified location.

### **Conclusion**

You may experience Excel file corruption warning messages while opening or saving an Excel file. The file may become corrupt due to malware infection, sudden system shutdown, and forgetting to close workbook on a shared network. This post outlined occurrences of Excel file corruption warnings, and also described solutions to fix the issue.

You may try using Microsoft’s built-in ‘Open and Repair’ tool to repair corrupt workbook and recover data from it. If this solution doesn’t work, proceed with uninstalling and re-installing the Office application. Another solution is to move corrupt files to another location. But if the problem still persists, use **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** software to repair single or multiple Excel (.xls or .xlsx) files and restore data.


## Excel AutoRecover not working, what is next?

## Consider a Scenario

A professional with Windows 10 computer had MS Office 2016 installed on it. For an official purpose, he worked on an Excel workbook and saved it as an XLSX file. After working for hours on it, which was saved with a file name, a power outage occurred in his building for quite some time. After the power was back, he reopened Excel to find a list of recovered files in ‘Document Recovery’ section on the screen’s left side. However, the file that he had worked on recently was the ‘Original version,’ i.e. the last version saved by him and not the auto-saved Excel file. This meant the Excel document did not have any new data that was entered since the last time he saved it. Consequently, he lost hours of work. According to him, this happened despite the fact that the ‘AutoRecover’ feature was enabled. (Still, this needs to be checked and ensured.)

The 'AutoRecover' feature might not work in any of these cases:

- **AutoRecover Feature is disabled -** With this feature disabled, the Excel files are not auto-saved if the document is closed without saving, or the document closes unexpectedly due to an untoward incidence. To check, see if ‘Save AutoRecover information every \* minutes’ and ‘Keep the last auto-saved version if I close without saving’ checkboxes are checked or unchecked. If either one is unchecked or both are unchecked, it signifies that the AutoRecover feature is disabled. Else, the AutoRecover is enabled.
- **Corruption in the Excel XLSX file –** If ‘AutoRecover’ is enabled, most probably the cause is ‘damaged Excel XLSX file.’

Before discussing solutions to resolve the ‘Excel AutoRecover not working’ issue, let's have an overview of the 'AutoRecover' and 'AutoSave' features.

## A Brief Overview of Excel AutoRecover and AutoSave Feature

AutoRecover is an inbuilt feature in MS Excel 2019, 2016, 2013, 2010, 2003, and 2007 that allows saving all of the ‘open Excel files’ at a fixed-interval in a user-specified location or AutoRecover file location. Besides, AutoSave is an add-in that exists in Excel 2002 and earlier versions to save all open Excel files in case of a crash, power outage, or accidental closure of Excel files without saving. Users can recover these files if Excel closes suddenly, for instance, at the time of power outage or failure. The auto-recovered Excel files are saved at a default location.

## Methods to Fix ‘Excel AutoRecover Not Working’ Issue

### Manual Methods

### Method 1 - Enable the 'AutoRecover' Feature if Disabled

Make sure that you have the ‘AutoRecover' feature enabled in your Excel application. If not, follow these steps to enable it:

- Open Excel with MS Excel 2016/2019
- Click on File and then on Options tab
- In ‘Excel Options’, click on Save tab
- Check ‘Save AutoRecover information every \* minutes’ and ‘Keep the last auto saved version if I close without saving’ box
- Set the time in ‘Save AutoRecover information every \* minutes’
- Click on the OK button

Note: With this method, it is not possible to recover data from the current Excel file. From next time onwards, the AutoRecover feature starts working following which Excel shall start auto-saving Excel files as per the time set in ‘Save AutoRecover information every \* minute’.

### Method 2 - Repair Corrupt Excel File

If corruption in Excel XLSX file has resulted in ‘Excel AutoRecover not working’ issue, you will need to [repair Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). Use **‘Open and Repair’** inbuilt utility to fix and repair the damaged (corrupt) Excel file and extract its data. To use the inbuilt utility, execute the following steps:

- Go to location “C:\\Users\\AppData\\Local\\Microsoft\\Office\\UnsavedFiles” to find Excel TMP files and save it as XLSX file

Note: In Windows 8, the location is the same as mentioned above that is for Windows 10. In Windows 7, the location is “C:\\Users\\name\\AppData\\Roaming\\Microsoft\\Excel\\”

- Open a blank **Excel** sheet; click **File >> Open**
- Go to the location and folder containing the damaged Excel file
- In the **Open** dialog box, choose the damaged Excel file and click the arrow next to the Open button, and then click **Open and Repair**
- In the window that appears, click **Repair** to recover as much data as possible

Now, open the Excel (XLSX) file to check if the Excel file is repaired and its data is recovered.

### Use a Professional Excel File Repair Tool

Using a third-party Excel file repair tool can help you repair damaged Excel XLSX file and recover all the data. **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** is one tool you can rely on to repair severely corrupt Excel files (XLSX or XLS).

Here are a few other reasons why you should choose Stellar Repair for Excel software:

- Repairs Excel file while keeping the worksheet properties and cell formatting same as before.
- Recovers all of the Excel file components like tables, forms, reports, charts, chart sheets, cell comments, formulas, images, etc.
- Can batch repair multiple Excel files simultaneously
- Supports Excel 2019 and earlier versions

[![Free download Stellar Repair for Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Conclusion

To help resolve the problem of ‘AutoRecover not working’, different methods have been discussed, depending on the cause of the problem. These solutions can be implemented to check the possibility of getting back maximum data added in last saved version of the Excel file. However, to fix corruption in the excel file, using a specialized tool such as Stellar Repair for Excel software recommended by MS Excel Experts and MVPs can help. The software can repair severely damaged Excel file easily and efficiently.


## Repair Office 2016 Files (Word, Excel and PowerPoint)on Windows

If you frequently work with Microsoft Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files, then issues like file inaccessibility or corruption won’t be new to you.

Let’s discuss some common scenarios which may lead to corrupt MS Office 2016 files:  

## Scenarios behind Microsoft Office Files Corruption

### Scenario 1 – Disruption during Data Migration

You decide to move Office files from your hard drive to other removable media. However, when you try to access the data within the files post-migration, you may find Word, Excel, and PowerPoint files showing gibberish characters. Due to a power surge, sudden system shutdown, and internal mechanical failure, the files may have turned corrupt.

![](https://www.stellarinfo.com/image/catalog/article/word/Word-displaying-gibberish-characters.png)

Figure 1- Microsoft Word file showing garbage characters

### Scenario 2 – Office Files and Registry Entries Become Infected

When you open or use the Microsoft Office application, it crashes as soon as it opens. You assume that an add-in was causing the problem and restart the Office application without add-ins loaded, but the application still crashes. This may happen because of a virus infecting the Office files and registry values, thus leading to corrupt or damaged Office files.

### **Scenario 3 – Inaccessible or Lost Data**

Suppose all your Office files are stored on a USB device, and you unplugged the device while it was still open in Windows. Now, when you attempt to open a Word or an Excel file, all the data is gone. Unsafe removal of USB or any other external storage device may corrupt the data inside your Office files or turn the file inaccessible.

## How Can You Deal with Microsoft Office Files Corruption?

Here are a few solutions that can help you fix or repair Office 2016 Files Corruption:

### Solution 1 – Use Microsoft in-built Repair Utility

Microsoft recommends using its in-built repair utility, 'Open and Repair', to fix corrupt Office files. Follow these steps to understand how you can use the utility to repair the corrupt Word, Excel, and PowerPoint files:

- Launch the MS Office application whose file you want to repair:

1. To repair corrupt Word (.doc, .docx) files, launch MS Word
2. To repair corrupt Excel files (.xls, .xlsx) files, launch MS Excel
3. To repair corrupt PowerPoint (.ppt, .pptx) files, launch MS PowerPoint

- Click File, and then click the Open tab.
- Click Navigate to the location or folder where the Word, Excel, or PowerPoint file is stored.
- Select the corrupt file you want to repair by single-clicking on it, and then find the Open button and click on the drop-down menu next to it.

![](https://www.stellarinfo.com/image/catalog/article/word/Open-and-repair.png)

- From the drop-down menu, click the **Open and Repair** option and follow the subsequent instructions to repair Office 2016 files.

### **Solution 2 – Repair Office 2016 Installation**

Try repairing the Office installation to fix the MS Office files. The steps to repair your Office installation may vary depending on the operating system you are using.

**For Windows 7**

- Open your PC's control panel
- Click **Programs**

![](https://www.stellarinfo.com/image/catalog/article/word/ControlPanel-Programs.png)

- Click **Programs and Features,** and then click **Uninstall a program** option

![](https://www.stellarinfo.com/image/catalog/article/word/ControlPanel-Uninstall.png)

- Right-click on the Office application you want to repair, and then click **Change**

![](https://www.stellarinfo.com/image/catalog/article/word/Change-Office-application.jpg)

- Under **Change your installation of Microsoft Office Professional Plus 2016,** choose Repair and then click **Continue.**

![](https://www.stellarinfo.com/image/catalog/article/word/Repair-Office-Application.jpg)

**For Windows 10**

- Right-click the Start button, and type in **Apps & Features** (For Windows 10)

![](https://www.stellarinfo.com/image/catalog/article/word/Apps-and-features.png)

**NOTE:** This step will work for Windows 10/8/8.1/7 and Vista

- Click **Programs** from the window that opens, click on the MS Office product you want to repair, and then click on **Modify**

![](https://www.stellarinfo.com/image/catalog/article/word/Modify-Office-application.jpg)

**Note:** Following the step will repair the entire Microsoft Office suite even if it contains only one application you want to repair such as an Excel or PowerPoint file. But, in case you have a standalone app installed, try to locate that application by name.

- Under **Change your installation of Microsoft Office Professional Plus 2016,** choose Repair, and then click **Continue** to initiate the repair process.

![](https://www.stellarinfo.com/image/catalog/article/word/Repair-Office-Application.jpg)

- Once the repair process completes, you'll be prompted to restart your PC. Click **Yes**

### **Solution 3 – Use Stellar Toolkit for File Repair**

Repair MS Office 2016 files by using [Stellar Toolkit for File Repair](https://tools.techidaily.com/stellardata-recovery/file-repair-toolkit/). This software comprises four essential utilities that can help you repair corrupt MS Word, MS Excel, MS PowerPoint, and PDF files.

The toolkit helps repair corrupt Office 2016 and other version documents and files while maintaining the original file format, which is less likely achievable with inbuilt methods. Follow these steps to repair MS Office 2016 documents by using the Office file repair tool:

- Download and install **Stellar Toolkit for File Repair**.

[![free download](https://www.stellarinfo.com/blog/wp-content/uploads/2021/11/free-download-1-4.png)](https://tools.techidaily.com/stellardata-recovery/file-repair-toolkit/)

- Launch the software.
- From the software's main interface, select the MS Office file you want to repair.

![](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Stellar-toolkit-for-file-repair-main-interface.png)

- From the window that pops up, select the corrupted file to be repaired.

Note: If you don't know the exact location of corrupt office files or if they are large in number, you can locate the files by using the Find/Search option included in the software.

- After selecting the file, click the Scan button to initiate the repairing process.
- Once the scanning process is complete, all the recoverable information is displayed in the software's left-hand panel. Click on any item to preview it before recovery.
- To save the repaired data, click the Save button, and enter a destination of your choice.
- Click OK.

## Conclusion

This post outlined possible scenarios and their causes that may lead to corruption in MS Office 2016 files. It also emphasized how the inbuilt methods such as Open and Repair, and Repair Office Installation help to resolve the corruption issues. But these are not competent enough to resolve all the errors. With Stellar Toolkit for File Repair, you can resolve all sorts of corruption issues and recover data of Office 2016 files – Excel, Word, PPT, and PDF – in their original state.


## How to fix “damage to the file was so extensive that repairs were not possible” Excel error?

**Summary:** Unable to resolve “damage to the file was so extensive that repair was not possible” error in Excel? Read this post to discover more details about the error, possible causes, and how to rectify the error. To save time & efforts, you can also try an Excel file repair software to resolve the “damage to the file…” error in a few clicks.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When opening a workbook in Microsoft Excel 2003 or later, you may encounter an error message,

_“Damage to the file was so extensive that repairs were not possible. Excel attempted to recover your formulas and values, but some data may have been lost or corrupted.”_

![](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-1.png)

The error message may also occur while exporting an Excel file. Let’s find out what causes this error and what we can do to fix it.

## Reasons Behind “Damage to the File Was So Extensive That Repairs Were Not Possible” Error

Your Excel file may be corrupt, oversized, virus-afflicted, etc., which can trigger this error and make the repair impossible. Below are some common reasons.

- Large or oversized excel files hindering export
- Data restore errors
- Field length of a cell is more than 256 characters
- Software conflicts, viruses, network failure
- Unable to open files in upgraded versions
- Errors on output exceeding 64000 rows
- Limited system resources (such as RAM, internal memory)

In a nutshell, the error generally happens if Excel discovers unreadable content, which may also interrupt file saving in Excel.

## How to Resolve “Damage to the File Was So Extensive That Repairs Were Not Possible” Error?

Here are a few methods you can follow to fix or resolve the Excel repair error.  

### Method 1: Perform Basic Troubleshooting

When opening a corrupt workbook, Microsoft Excel automatically initiates the file recovery mode to repair the corrupt file. However, if it fails to perform automatic recovery, then follow these basic troubleshooting steps:

- This error mainly happens when you try to open the Excel file in an upgraded version.  Try to open the file in an older version of Excel. You might be able to open it.
- Try saving the file with a different file name.
- Use a different file extension to save the file.
- You can save the Excel file as HTML and then open it. However, an HTML file might not save conditional formatting.
- Close other opened applications on the system which may be causing the error.
- Select less data for export at once.
- Delete worksheets if copied from another document; for instance, delete any file or screenshots you have imported.
- Open the file on another system.

If the error persists, then use the manual method to repair a workbook using the below steps:

- Go to the “**File**” tab.

![Go the File Tab in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-2.png)

- Select **Open** and select the damaged spreadsheet from the **Recent Workbooks** section on the right, if listed. However, if you cannot find the file in the **Recent Workbooks** section, click on “**Browse**” and choose the **corrupted workbook**.

![Click on Open and browse the corrupted workbook](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-3.png)

- Click the drop-down arrow on the **Open** tab and select **Open and Repair**.

![after selecting the corrupt excel, click on the drop-down next to Open and click on Open and Repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-4.png)

### Method 2: Check if exporting a Heavy File is Causing Resource Limitations in Excel

Sometimes, when you try to export an Excel sheet carrying a huge database, you may face memory errors in older Excel versions like Excel 2003. Here, you’ll have to decrease the amount of data as Excel 2003 does not permit exporting extensive data beyond a limit. However, modern versions such as Excel 2007, 2010 & 2016 allow exporting a large amount of data and utilize more RAM than the older versions.

Following are some other workarounds:

- Use a lesser number of query presentation fields to re-generate the query. Then, again re-enter those fields.
- Decrease the multi-line string field data text up to 8000 characters.

### Method 3: Copy Macros and Data to Another Workbook (Empty) in an Advanced Excel version

If the issue is occurring due to version incompatibility, i.e., if the file opens easily in the older version but shows errors in the new version. You can:

**Use the older version to open the file or copy the data or macros in an empty workbook of the new version of Excel.**  

#### Copying the Macros in the Workstation

In Microsoft Excel, you can use the **Visual Basic Editor** to open the workbook with macro on another workbook by copying the macro. Both VBA tools and Macros appear in the Developer section of the excel file. This option is disabled by default. So first, you need to enable it.

Follow the instructions to enable it:

- Open Excel and go to **File** > **Options.**

![In the file menu, go to Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-5.png)

- Click “**Customize Ribbon.**”

![In Excel options, click on Customize Ribbon](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-6.png)

- Look at the right side of the pane and ensure the **Developer** tab is checked.

![At the right side, make sure that the "Developer" box is checked](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-7.png)

- Click **OK**.

Once you have enabled the **Developer** tab, follow the steps to copy the macro from one workbook to another:

- First, open both the workbooks- the workbook containing the macro and the workbook in which you need to copy the macros.
- Locate the **Developer** tab.

![Developer tab in Ribbon](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-8.png)

- Select **Visual Basic** to display the “**Visual Basic Editor**.”  

![Click on Visual Basic in the Developer Tab](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-9.png)

- Go to the View menu in the Visual Basic Editor.

![Go to the View menu in the Visual Basic Editor](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-10.png)

- Select **Project Explorer**.

![Click on Project Explorer from the View Menu](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-11.png)

- In the **Project Explorer** window, drag the module you need to copy to the destination workbook. For example:

![In project explorer, Drag the module you need to copy to the destination workbook](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-12.png)

**Module 1** has been copied from **Book2.xlsm** to **Book1.xlsm**

### Method 4- Restore the backup file

The workbook backup helps to open the corrupted or mistakenly deleted file. Sometimes, the issue can be fixed using the **Recover Unsaved Workbook** option in Excel. Here’s the list of steps to recover the files in Microsoft Excel:

- Go to the **File** tab on Excel.

![Go to the File Menu](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-13.png)

- Click **Open**.

![Click on Open ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-14.png)

- Search on the top-left of the screen to click **Recent Workbooks** as below:

![Click on Recent ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-15.png)

- Next, scroll down to the bottom.
- Click the “**Recover unsaved workbooks**” button.

![At the bottom, click on Recover unsaved workbooks](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/07/image-16.png)

- Scroll and find the lost file.
- Now double-click on the **file** to open.

[![Stellar Repair for Excel](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

### Conclusion

“**Damage to the file was so extensive that repairs were not possible**” error can be fixed with the above troubleshooting methods or by using a third-party Excel repair tool, like **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)**. Although There are no standard resolutions to fix the excel error as they may vary with different scenarios. In some cases, the manual methods might be time-consuming or fail to fix the error or recover the excel file. Hence, using an excel file repair tool may be the best option! It extracts data from the corrupted file and saves it to a new Excel workbook, which you can open and edit.


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
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-lava-blaze-2-5g-by-fonelab-android-recover-video/"><u>How to recover old videos from your Lava Blaze 2 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-poco-m6-pro-5g-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Poco M6 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-15-pro-max-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 15 Pro Max?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-12-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 12 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-6s-data-from-itunes-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 6s Data From iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-drivers-with-windows-device-manager-on-windows-10-by-drivereasy-guide/"><u>How to identify missing drivers with Windows Device Manager on Windows 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-11-pro-max-ios-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 11 Pro Max iOS System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-y100-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Vivo Y100 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-xs-max-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone XS Max | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-x50-by-fonelab-android-recover-photos/"><u>How to get back lost photos from X50.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-lava-blaze-2-by-fonelab-android-recover-video/"><u>How to recover old videos from your Lava Blaze 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-oppo-find-x7-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Oppo Find X7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-honor-x50-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Honor X50 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-vivo-y77t-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Vivo Y77t?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-vivo-s17-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Vivo S17</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-tecno-spark-10-5g-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Tecno Spark 10 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-contacts-on-iphone-7-plus-4-methods-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore Contacts on iPhone 7 Plus (4 Methods) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-x-to-an-older-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone X to an Older Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-multiple-pdf-files-by-stellar-guide/"><u>How to Repair Multiple PDF Files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-motorola-razr-40-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Motorola Razr 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-nubia-red-magic-9-pro-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Nubia Red Magic 9 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-docm-file-document-with-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .docm file document with Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-realme-11x-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Realme 11X 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-oppo-a1-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Oppo A1 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h265-video-on-xiaomi-redmi-note-12t-pro-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Xiaomi Redmi Note 12T Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-samsung-galaxy-a23-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Samsung Galaxy A23 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-huawei-nova-y91-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Huawei Nova Y91?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-iphone-6s-stuck-at-attempting-data-recovery-loop-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix iPhone 6s Stuck at attempting data recovery Loop | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-ios-system-issues-of-iphone-14-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iOS System Issues of iPhone 14? | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-best-8-tools-tell-you-how-to-make-slideshows-for-instagram/"><u>Updated Best 8 Tools Tell You How to Make Slideshows for Instagram</u></a></li>
<li><a href="https://howto.techidaily.com/9-solutions-to-fix-process-system-isnt-responding-error-on-nokia-c210-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>9 Solutions to Fix Process System Isnt Responding Error on Nokia C210 | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-easy-shortcut-to-convert-videos-to-gifs-on-iphone/"><u>New Easy Shortcut to Convert Videos to Gifs on iPhone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/full-guide-to-bypass-vivo-v30-pro-frp-by-drfone-android/"><u>Full Guide to Bypass Vivo V30 Pro FRP</u></a></li>
<li><a href="https://techidaily.com/how-to-factory-reset-oneplus-11r-in-5-easy-ways-drfone-by-drfone-reset-android-reset-android/"><u>How to Factory Reset OnePlus 11R in 5 Easy Ways | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/does-13-ultra-has-native-mov-support-by-aiseesoft-video-converter-play-mov-on-android/"><u>Does 13 Ultra has native MOV support?</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/can-i-bypass-a-forgotten-phone-password-of-realme-narzo-60x-5g-by-drfone-android/"><u>Can I Bypass a Forgotten Phone Password Of Realme Narzo 60x 5G?</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-5-ways-to-transfer-music-from-honor-90-gt-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 5 Ways to Transfer Music from Honor 90 GT to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/5-ways-to-reset-tecno-spark-10c-without-volume-buttons-drfone-by-drfone-reset-android-reset-android/"><u>5 Ways to Reset Tecno Spark 10C Without Volume Buttons | Dr.fone</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-fix-unresponsive-phone-touchscreen-of-xiaomi-civi-3-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How To Fix Unresponsive Phone Touchscreen Of Xiaomi Civi 3 | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/updated-best-8-online-gif-to-apng-converters-for-2024/"><u>Updated Best 8 Online GIF to APNG Converters for 2024</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-fix-apple-id-verification-code-not-working-on-apple-iphone-12-by-drfone-ios/"><u>How To Fix Apple ID Verification Code Not Working On Apple iPhone 12</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-planning-to-use-a-pokemon-go-joystick-on-apple-iphone-11-pro-max-drfone-by-drfone-virtual-ios/"><u>In 2024, Planning to Use a Pokemon Go Joystick on Apple iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/where-is-the-best-place-to-catch-dratini-on-apple-iphone-12-drfone-by-drfone-virtual-ios/"><u>Where Is the Best Place to Catch Dratini On Apple iPhone 12 | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/in-2024-11-best-location-changers-for-honor-magic-6-drfone-by-drfone-virtual-android/"><u>In 2024, 11 Best Location Changers for Honor Magic 6 | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-get-the-apple-id-verification-code-on-apple-iphone-xs-in-the-best-ways-by-drfone-ios/"><u>How To Get the Apple ID Verification Code On Apple iPhone XS in the Best Ways</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-activate-and-use-life360-ghost-mode-on-nokia-105-classic-drfone-by-drfone-virtual-android/"><u>In 2024, How To Activate and Use Life360 Ghost Mode On Nokia 105 Classic | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-video-files-on-htc-u23-pro-by-fonelab-android-recover-video/"><u>Complete guide for recovering video files on HTC U23 Pro</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-8-ways-to-transfer-photos-from-oppo-find-x7-to-iphone-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 8 Ways to Transfer Photos from Oppo Find X7 to iPhone Easily | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/4k-video-editing-made-easy-top-proxy-editing-tools-for-2024/"><u>4K Video Editing Made Easy Top Proxy Editing Tools for 2024</u></a></li>
<li><a href="https://change-location.techidaily.com/how-does-the-stardust-trade-cost-in-pokemon-go-on-samsung-galaxy-a54-5g-drfone-by-drfone-virtual-android/"><u>How does the stardust trade cost In pokemon go On Samsung Galaxy A54 5G? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-best-tools-to-hard-reset-huawei-nova-y91-drfone-by-drfone-reset-android-reset-android/"><u>3 Best Tools to Hard Reset Huawei Nova Y91 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/failed-to-play-hevc-files-on-galaxy-a23-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>Failed to play HEVC files on Galaxy A23 5G</u></a></li>
<li><a href="https://android-location.techidaily.com/getting-the-pokemon-go-gps-signal-not-found-11-error-in-google-pixel-7a-drfone-by-drfone-virtual/"><u>Getting the Pokemon Go GPS Signal Not Found 11 Error in Google Pixel 7a | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-how-to-unlock-poco-m6-pro-5g-phone-pattern-lock-without-factory-reset-by-drfone-android/"><u>In 2024, How to Unlock Poco M6 Pro 5G Phone Pattern Lock without Factory Reset</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-data-from-nokia-c12-pro-to-other-android-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Data from Nokia C12 Pro to Other Android Devices? | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/ways-to-fix-the-failed-to-parse-the-corrupted-excel-2013-file-error-by-stellar-guide/"><u>Ways to Fix the Failed to Parse the Corrupted Excel 2013 File Error</u></a></li>
</ul></div>



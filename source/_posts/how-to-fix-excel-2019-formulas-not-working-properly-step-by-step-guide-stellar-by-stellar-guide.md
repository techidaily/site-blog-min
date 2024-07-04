---
title: How to Fix Excel 2019 Formulas Not Working Properly | Step-by-Step Guide | Stellar
date: 2024-05-19T18:32:11.553Z
updated: 2024-05-20T18:32:11.553Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2019 Formulas Not Working Properly | Step-by-Step Guide
excerpt: This article describes How to Fix Excel 2019 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair excel 2000,repair damaged .xltm,repair .xltx files,repair .xlb files,repair excel 2003,repair excel,repair corrupt .xls files
thumbnail: https://www.lifewire.com/thmb/zWxNYaUUEWKWwcCxwdQ9gx8_XCw=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/how-to-safely-use-airport-charging-stations-4690583-01-6085c56e46b34f84b5f5d2d976e20677.jpg
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


## Top 5 Ways to Fix Excel File Not Opening Error

**Summary:** MS Excel users sometimes face issues while using the MS Excel application. One such issue is the Excel file not opening error. In this post, we’ve mentioned the reasons that may result in this error and the ways to resolve it. Also, you’ll find about an Excel repair software that can help you repair corrupt Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Several Microsoft Excel users have reported encountering the ‘Excel file not opening’ error when opening their Excel file. There are several reasons that may cause this error. In this post, we’ll be discussing the reasons that may lead to the ‘Excel file not opening’ error and the top 5 ways to fix this error.

## Why Does the ‘Excel File Not Opening’ Error Appear?

Following are some possible causes that may result in the ‘Excel file not opening’ error:

1. There may be a problem with an add-in that is preventing you from opening the Excel files.
2. There’s a chance that your Excel application is faulty.
3. Your Excel program is unable to communicate with other programs or the operating system.
4. The file association might have been broken. This is a common problem faced by users who have upgraded their Excel application or operating system.
5. The file you’re trying to open is corrupted.

## 5 Ways to Fix Excel File Not Opening Error

Let’s explore the ways to resolve the Excel file not opening error:

### 1\. Uncheck the Ignore DDE Checkbox

Dynamic Data Exchange (DDE)allows Excel to communicate with other programs. The Excel error may occur due to incorrect DDE settings. You need to ensure that the correct DDE configuration is enabled. Follow the steps provided below:

- Launch your MS Excel file.
- Go to **File > Options**.

![options menu](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/source/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/options-menu-01.png)

- Now click on **Advanced**.

![Advanced options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/Advanced-options-02.png)

- Further, find the **General** option on the screen.

![General Option uncheck Dynamic Data Exchange Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/general-option-uncheck-Dynamic-Data-Exchange-option-03.png)

- Uncheck the option **‘Ignore other applications that use Dynamic Data Exchange (DDE)**’.
- Click **OK** to save the changes.

### 2\. Reset Excel File Associations

When you launch your Excel file, the file association ensures that the Excel application is used to open the file. You can try to reset these associations and see if Excel opens after this. Proceed with the following steps to do so:

- Navigate to **Start Menu** and launch **Control Panel**.
- Now, navigate to **Programs > Default Programs > Set Your Default Programs**.

![Set-Your Default Programs](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/set-your-default-programs-04.png)

- A new window will open. Herein, find the Excel program in the list and select it. Now, select the option **‘Choose defaults for this program’**. Click **OK**.

![Set Defaults](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/set-defaults-05.png)

- A new window for **‘Set Program Associations’** will open.
- Check the box against the **‘Select All’** option.
- Further, click **Save** to reset the Excel File Associations settings.

![Set Program Associations](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/set-program-associations-06.png)

### 3\. Disable Add-Ins

Many people install third-party add-ins to enhance the application’s functionality. Sometimes, these add-ins can create an issue. Follow the below-mentioned steps to disable the problem creating add-ins:

- Launch MS Excel application.
- Navigate to **File > Options > Add-ins**.

![Click on Add ins Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/Click-on-Add-ins-option-07.png)

- In the window that opens, go to the **Manage** option at the bottom.
- Herein, select the **COM Add-ins** option from the dropdown list. Click **Go.**

![Select Com Addins](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/Select-Com-Addins-08.png)

- In the COM Add-ins window, uncheck all the boxes to disable the add-ins. Click **OK.**

![Diasble Com Add ins Checkbox](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/diasble-com-add-ins-checkbox-09.png)

### 4\. Repair MS Office Program

Sometimes the issue is not with your Excel file. Instead, the reason for the error can be a corrupt MS Office application. You can repair the program to fix the Excel file not opening error. Here are the steps:

- Press the **Windows + R** keys to launch the **‘Run’** dialog box.

![Run Dialog box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/Run-Dialog-box-10.png)

- Enter the text **‘appwiz.cpl’** to launch the program and features window.

![appwiz.cpl in the run field](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/appwiz.cpl-in-the-run-field-11.png)

- Find the MS Office program in the list of applications.

![locate Microsoft Office in the list](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/locate-Microsoft-Office-in-the-list-12.png)

- Right-click on it and select **Change**.

![Right Click and Select Change](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/right-click-and-select-change-13.png)

- In the new window, select the **Quick Repair** radio button. Click **Repair**.

![Quick Repair Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/quick-repair-option-14.png)

- Follow the on-screen instructions to repair the Office application. Once the repair process is completed, you can try opening the Excel file to see if the problem is resolved.

### 5\. Disable Hardware Graphics Acceleration

The hardware graphics acceleration assists in the system’s better performance, especially when you use MS Office applications, like MS Excel or Word. Sometimes, this causes the Excel file not opening issue. You can disable this option to try to resolve the issue. Here are the steps:

- Launch your MS Excel application.
- Navigate to **File > Options > Advanced**.
- Herein, go to the **Display** option.
- Uncheck the **Disable hardware graphics acceleration** checkbox. Click OK.

![Uncheck Disable Hardware Acceleration](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/10/hardware-accelaration-15.png)

## What If These Solutions Do Not Work?

If you have applied all the methods mentioned above and still cannot open your Excel file, there are chances that your file is corrupted. You can use a specialized [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel to repair the corrupted Excel file. This software has powerful algorithms that can scan and repair even severely corrupt Excel files, without any file size limitation. After repairing the file, it restores all the data, including tables, charts, rules, etc. to a new Excel, with 100% integrity.

To know how the software works, see the video below:

<iframe title="How to Repair and Recover Corrupted Excel Files?" width="750" height="422" frameborder="0" allowfullscreen="" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" nitro-og-src="https://www.youtube.com/embed/3SiJqmP7iMU?feature=oembed&amp;autoplay=1" nitro-lazy-src="data:text/html;https://www.youtube.com/embed/3SiJqmP7iMU?feature=oembed&amp;autoplay=1;base64,PGJvZHkgc3R5bGU9J3dpZHRoOjEwMCU7aGVpZ2h0OjEwMCU7bWFyZ2luOjA7cGFkZGluZzowO2JhY2tncm91bmQ6dXJsKGh0dHBzOi8vaW1nLnlvdXR1YmUuY29tL3ZpLzNTaUpxbVA3aU1VLzAuanBnKSBjZW50ZXIvMTAwJSBuby1yZXBlYXQnPjxzdHlsZT5ib2R5ey0tYnRuQmFja2dyb3VuZDpyZ2JhKDAsMCwwLC42NSk7fWJvZHk6aG92ZXJ7LS1idG5CYWNrZ3JvdW5kOnJnYmEoMCwwLDApO2N1cnNvcjpwb2ludGVyO30jcGxheUJ0bntkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y2xlYXI6Ym90aDt3aWR0aDoxMDBweDtoZWlnaHQ6NzBweDtsaW5lLWhlaWdodDo3MHB4O2ZvbnQtc2l6ZTo0NXB4O2JhY2tncm91bmQ6dmFyKC0tYnRuQmFja2dyb3VuZCk7dGV4dC1hbGlnbjpjZW50ZXI7Y29sb3I6I2ZmZjtib3JkZXItcmFkaXVzOjE4cHg7dmVydGljYWwtYWxpZ246bWlkZGxlO3Bvc2l0aW9uOmFic29sdXRlO3RvcDo1MCU7bGVmdDo1MCU7bWFyZ2luLWxlZnQ6LTUwcHg7bWFyZ2luLXRvcDotMzVweH0jcGxheUFycm93e3dpZHRoOjA7aGVpZ2h0OjA7Ym9yZGVyLXRvcDoxNXB4IHNvbGlkIHRyYW5zcGFyZW50O2JvcmRlci1ib3R0b206MTVweCBzb2xpZCB0cmFuc3BhcmVudDtib3JkZXItbGVmdDoyNXB4IHNvbGlkICNmZmY7fTwvc3R5bGU+PGRpdiBpZD0ncGxheUJ0bic+PGRpdiBpZD0ncGxheUFycm93Jz48L2Rpdj48L2Rpdj48c2NyaXB0PmRvY3VtZW50LmJvZHkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBmdW5jdGlvbigpe3dpbmRvdy5wYXJlbnQucG9zdE1lc3NhZ2Uoe2FjdGlvbjogJ3BsYXlCdG5DbGlja2VkJ30sICcqJyk7fSk7PC9zY3JpcHQ+PC9ib2R5Pg=="></iframe>

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Conclusion

Before you proceed with resolving the Excel file not opening error, try to find out the root cause of this error. If you know the real reason, you can try the method right away. If the reason for the error is corruption in the Excel file, the best option is to repair the file using a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).


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


**Summary:** Recovering your unsaved Excel file on Windows 10 with some of our tested methods can truly save you a lot of time, money, and effort. Here’s how to do it!

Microsoft provides a wide range of applications to its users that cater to different purposes. Out of all the MS Office applications, Excel is considered the most used native application by several companies and individuals to store and process crucial data. Excel can also be called a go-to data analysis tool that offers great reliability, ease of use, and multiple features. But unfortunately, it’s not immune to human errors and system errors.

What if you accidentally close the Excel file without saving your work or Excel crashes suddenly? Or worse! Your system crashes without any warning and when you reboot your system, you can’t find the unsaved Excel file.

There could be different situations where you may lose access to your unsaved Excel workbook. There is no need to be anxious, though! We’ve covered some easy-to-follow DIY methods in this post to help you recover unsaved Excel files quickly.

## **4 Ways to Recover Unsaved Excel File**

Irrespective of the circumstances, you can recover unsaved Excel files on Windows 10 using the following methods.

## **Method 1: Use the Search Option**

If you remember the name of the Excel file, try to find it using the Search option. Following these steps:

- Go to **Windows Search Box** and type the name of the file.
- If you find the file there, click **Open** or **Open file location**.

![choose-open-or-open-file-location-from-Windows-search-box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/Type-excel-file-name-in-windows-search-box-and-click-open-or-open-file-location-image1.jpg)

- If you click **Open file location**, it will take you to the exact location of the file, and you’ll be able to regain access to lost/unsaved Excel file easily.

## **Method 2: Use AutoRecover Feature**

Microsoft Excel comes with a built-in AutoRecover feature that saves copies of all open Excel files at a user-definable fixed interval. You can recover the file if you forget to save it and accidentally close it or it closes automatically due to an unexpected system breakdown or power failure. You can follow the given steps to recover unsaved Excel files with the AutoRecover feature:

- Open a new Excel file and go to **File > Options**.
- Then go to **Save > Save** **workbooks**.
- Next, ensure that the ‘**Save AutoRecover information every’** and ‘**Keep the last autosaved version if I close without saving**’ option is already selected.

![check-if-the-autorecover-options-for-excel-files-are-selected ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/check-if-the-autorecover-options-for-excel-files-are-selected-image2.jpg)

- Now, copy the file path given against the **AutoRecover file location**.
- Open the **File Explorer**, paste the Menu Bar file path there, and hit **Enter**.
- The unsaved file will appear with the **.xlb extension**.
- Double-click the file and choose **Excel** from apps to open the **.xlb file**.

![choose-excel-application-to-open-xlb-file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/choose-excel-application-to-open-xlb-file-image4.jpg)

- Finally, save the file to the desired location.

## **Method 3: Use the ‘Recover Unsaved Workbooks’ Option**

Another way that you can choose is the ‘**Recover Unsaved Workbooks’** option. To regain access to your Excel file, follow the given steps:

- Open an **Excel** sheet and then go to **File > Open**.
- Now, scroll down and click **Recover Unsaved Workbooks**.

![go-to-file-then-open-and-click-recover-unsaved-workbooks](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/go-to-file-then-open-and-click-recover-unsaved-workbooks-image5-1024x674.jpg)

- The list of unsaved files will show in the **Document Recovery** pane.
- Locate the file you lost and double-click on it to open.
- Finally, right-click on the version of the file you want to restore and click **Save As**.

![unsaved-files-will-show-in-document-recovery-pane](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/unsaved-files-will-show-in-document-recovery-pane-image6.jpg)

And, save the file to the desired location

## **Method 4: Restore File with Previous Versions Utility**

This feature will help you recover unsaved Excel files when you’ve saved the workbook earlier, but you end up losing access to the file due to an Excel or system crash. It’d help you save the previously saved version of your Excel file. To do so, follow the given steps:

- Open an Excel file and navigate to **File > Info > Manage Versions**.

![go-to-file-then-info-and-then-manage-versions](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/01/go-to-file-then-info-and-then-manage-versions-image7.jpg)

- It will show the previously saved versions. You can choose the one you want to restore and click **Save/Save As** to store the file with a different name.

Hopefully, this blog has helped you learn how to recover unsaved Excel files. You can try any of the above methods.

## **FAQ**

**1. How can I repair a corrupted workbook?**

Whenever Excel detects a corrupted workbook while opening, it automatically starts running the File Recovery mode and repairing the file. If the File Recovery mode doesn’t start automatically, you can manually [repair the corrupted workbook](https://support.microsoft.com/en-us/office/repair-a-corrupted-workbook-153a45f4-6cab-44b1-93ca-801ddcd4ea53).




## \[Solved\] Excel Spreadsheet Disappears after Opening

Several Excel users have reported about experiencing ‘Excel spreadsheet disappears after opening’ issue. The problem occurs when attempting to open an Excel file by double-clicking on the file icon or name. The Excel file opens blank grey screen.

![Excel blank screen](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img1.jpg)

Figure 1 - Excel Blank Screen

## **User Instances of ‘Excel Open But Can’t See Spreadsheet’ Issue**

**Instance 1:** The user said that _“Excel 2016 opens to a blank screen and everything is greyed out. Ribbons at the top of Excel are also not present”._

**Instance 2:** The user reported that _“Attempting to open a spreadsheet, either by double-clicking the file or by opening Excel, it opens but only a blank Excel window is visible. When trying to save that particular file, Excel behaves as though there is no file open at all and all the file saving options, such as save, save as, print, etc. are greyed out.”_

## **Causes Behind ‘Excel Open But Can’t See Spreadsheet’ Issue & the Solutions Thereof**

Following are some of the possible reasons behind the ‘Excel open but can’t see spreadsheet’ issue, along with their solutions:

### **Cause 1: Excel File is Hidden**

You may have saved the Excel sheet as a hidden document.

### **Solution: Unhide Excel File**

Verify if your Excel worksheet is hidden by following these steps:

- In Excel, click the **View** tab, and then click **Unhide**.

Note: If the ‘Hide’ tab under the View menu is greyed out, the sheet you’re trying to open is not hidden. In that case, proceed to the next workaround.

![unhide excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img2.jpg)

Figure 2 - Unhide Excel File

- A dialog box will pop-up showing the worksheet name that can be unhidden. For instance, below is an image of Unhide box that lists ‘Daily\_Reports’ spreadsheet under Unhide workbook.

![unhide excel workbook](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img3.jpg)

Figure 3 – Select and Unhide Excel Workbook

- Click **OK** to unhide your Excel sheet.

If this fails to work, there’s a possibility that the spreadsheet window pane may have slided to one side of the visible desktop. To bring back the displayed area, click the ‘Arrange All’ option under the View tab. If the issue persists, try the next solution.

### **Cause 2: Ignore Dynamic Data Exchange (DDE) Option is Enabled**

Another reason behind the ‘Excel opens to a blank screen’ issue could be that the “Ignore other applications that use Dynamic Data Exchange (DDE)” checkbox is checked in Excel options.

Usually, when you double-click an Excel file, a DDE message is sent to Excel instructing it to open that particular file. But, if you have the ‘Ignore DDE’ option selected, the DDE message sent to Excel to open a workbook is ignored. As a result, the Excel workbook opens a blank screen.

### **Solution: Uncheck the Ignore DDE Option**

Make sure that the ‘Ignore DDE’ option is unchecked by performing these steps:

Note: Skip these steps for Excel 2019 and Excel Office 365.

- In your Excel window, click **File** > **Options**.

![Excel options](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img4.jpg)

Figure 4 - Select Excel Options

- From the left-side of the ‘Options’ window, choose **Advanced**.

![Excel Options Window](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img5.jpg)

Figure 5 - Excel Options Window

- In ‘Advanced’ window, locate the **General** section, and then uncheck the “Ignore other applications that use Dynamic Data Exchange (DDE)” checkbox.

![uncheck ignore DDE](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img6.jpg)

Figure 6 - Uncheck Ignore Other Applications that use DDE

- Click **OK** to apply the changes.

If these steps don’t work, proceed to the next workaround.

### **Cause 3: Issue within Office Program**

Sometimes, problem within your Microsoft Office program might cause Excel to behave oddly and may result in an Excel spreadsheet disappeared issue.

### **Solution: Repair Office Program**

Try repairing your Office program by executing the steps listed below, based on your Windows OS:

**For Windows 10:**

- Type Settings in the Windows search box.
- Click **Settings**.
- In ‘Windows Settings’ screen, select **Apps**.

![Windows settings screen](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img7.jpg)

Figure 7 - Windows Settings Screen

- In ‘Apps & features’ screen, scroll down to your Microsoft Office program and tap on it, and then click the **Modify**

![modify MS Office program](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img8.jpg)

Figure 8 - Modify MS Office Program

- In ‘How would you like to repair your Office programs’ dialog box, select the **Online Repair**radio button, and then click the **Repair**

![repair MS Office program](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img9.jpg)

Figure 9 - Repair MS Office Program

**For Windows 8:**

- Open Control Panel and click the **Uninstall a program** option under **Programs**.
- Select **Microsoft Office 365** and then click **Change**.
- In the window that opens, select **Online Repair** and then hit the **Repair**

You may be asked to restart your system after completion of the repair process.

**For Windows 7:**

- In ‘Control Panel’ window, double-click **Programs and Features**.
- Under **Uninstall or change a program** section, select the Office program, and then select **Change**.
- In the dialog box that appears, choose **Online Repair** and then choose **Repair**.

If repairing your Office installation fails to resolve the issue, continue to solution 4.

### **Cause 4: Problematic Excel and COM Add-ins**

Sometimes, Excel and COM add-ins may cause Excel file open a blank grey screen problem.

### **Solution: Disable the Add-ins**

Disabling the Excel and COM add-ins one at a time may help resolve the problem. The steps are as follows:

- Click the **File** menu in your Excel file and select
- From ‘Excel Options’ left-side panel, click **Add-Ins**, and then choose **COM Add-ins** from the ‘Manage’ drop-down. Click the **Go**

![COM Add-ins](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img10.jpg)

Figure 10 - Select COM Add-ins

- Uncheck one of the add-ins checkbox from the COM Add-ins window, then select **OK**.

![disable COM Add-ins](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img11.jpg)

Figure 11 - Uncheck and Disable COM Add-ins

- Restart your Excel program.

If the issue persists, repeat the above steps (1 till 4), except that you need to choose a different add-in in step 3.

If you’re still experiencing the same problem after unchecking all the COM add-ins, repeat all the above steps, except choose ‘Excel Add-ins’ in step 2.

If Excel can load the file, then the add-in that you disabled last leads to the problem. If turning off add-ins does not work, try the next solution.

### **Cause 5: Problem with Excel File Associations**

The Excel file opening a blank screen problem may occur if the file associations are not performing correctly.

### **Solution: Reset Excel File Associations**

Try resetting the file associations in Excel to their default settings. For this, follow the below steps based on your OS.

**For Windows 10 and Windows 8.1:**

- Select the file that is opening incorrectly and copy it to the desktop.
- Right-click the file and click **Properties**.
- In the File's 'Properties' window, check your file type next to **Type of File**from the General tab. For instance, (.docx), (.csv), or (.pdf).
- Next, check to which app your file is associated with from the **Opens with**option

If the file type is different than .xlsx, open the file in a different application by following these steps:

- Click the **Change button next to the ‘Opens with’ option.**
- Click **More apps**.
- From the list of applications, select the desired app and then check the **Always use this app** Click **OK**.

**For Windows 8:**

- Open **Control Panel**, click **Default Programs**, and then select **Set your default programs**.
- Choose **Excel** and then tap the **Choose default for this program**
- In ‘**Set Program Associations**’ window, choose **Select All** and then click **Save**.

**For Windows 7:**

- In **Control Panel**, choose **Default Programs**.
- Click the **Associate a file type or protocol with a specific program**
- Choose **Microsoft Excel Worksheet** and click on change program.
- Choose **Microsoft Excel** under **Recommended Programs**.
- If you’re unable to find Excel, browse the Excel installation folder.
- Select **exe** and then choose **Excel**.

## **What Next? Use Stellar Repair for Excel to Recover Your File**

If the above solutions don’t help recover your Excel file, use an Excel file repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to recover it. The software can restore the Excel spreadsheet and its components, including tables, pivot tables, charts, formulas, etc. while preserving the spreadsheet properties and formatting.

[![Free download Stellar Repair for Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Check out the video to recover Excel file by using Stellar Repair for Excel software.

<iframe src="//www.youtube.com/embed/VAeGzHnETu0" width="640" height="360" frameborder="0"></iframe>

## **Conclusion**

This article described the possible causes behind the ‘Excel open but can’t see spreadsheet’ issue, along with their solutions. The manual solutions to resolve the issue require time and efforts. But, if you need to access your Excel file without any delay, use Stellar Repair for Excel software to recover your .xlsx/.xls file in just a few clicks.


## Repair Files using Stellar Toolkit for File Repair

<a href="https://secure.2checkout.com/order/cart.php?PRODS=38733153&QTY=1&AFFILIATE=108875">Stellar Toolkit for File Repair Technician</a>

The main interface of Stellar Toolkit for File Repair comprises four modules to repair MS Office and PDF files. These modules are:

- Repair Document
- Repair Spreadsheet
- Repair PowerPoint
- Repair PDF

Click on the desired tab to repair that file format.

![Homepage of Stellar Toolkit for File Repair](https://www.stellarinfo.com/screenshots/file-toolkit/home-screen.png)

                                    _<small>Figure 1 - Illustrates Homepage of the Stellar Toolkit for File Repair</small>_

**Steps to Repair MS Word – .doc/.docx file**

- Click **Select File** to select a single corrupt Word (.doc/.docx) file that you want to repair. Alternately, click **Select Folder** for selecting all Word files in a single folder.

**_Note:_** _Click Find file(s) to search for the Word file, if the location is not known._

![Select word file](https://stellarinfo.com/support/kb/images/Select-word-file.jpg)

                                     _<small>Figure 2 - Illustrates Selection of single doc/.docx file or multiple files</small>_

- Once the file is selected, click the **Scan** button to scan and repair the file.
- A preview of the repaired Word file is displayed on the screen. Verify the file contents from the right pane of the preview window.

![Preview of word repair](https://stellarinfo.com/support/kb/images/preview-repaired-word-file.png)

                                         _<small>Figure 3 - Preview of Repaired Word Document</small>_

**_Note:_** _If you’re unable to repair a corrupt .doc file, select ‘Advance Repair’ option from the File menu for repairing the .doc files._  

- Click the **Save** icon on the **File** menu to save the repaired file.

![Select menu](https://stellarinfo.com/support/kb/images/file-menu.png)

                                                                     _<small>Figure 4 - File Menu</small>_

- In **Save Document** dialog box that appears, do the following:

- Select default location or a new folder to save the repaired file.
- Save the file in any of these formats: 'Full Document', 'Filtered Text' or 'Raw Text'.
- Click **OK**.

![saving word document](https://stellarinfo.com/support/kb/images/word-document-saving-option.png)

                                                        _<small>Figure 5 - Word Document Saving Options</small>_

The repaired file will be saved at your preferred location.

**Steps to Repair Excel – .xls/.xlsx files**

- In **Select File** window, click **Browse** to select the corrupt Excel file from the desired location. If you do not know the file location, click **Search** to find and select the corrupted spreadsheet.
- Once the Excel file is selected, start repairing the file by clicking the **Repair** button.

![Select xls/xlsx file](https://www.stellarinfo.com/screenshots/excel-repair/excel-window/2.jpg)

                              _<small>Figure 6 - Illustrates selection of one xls/xlsx file or multiple files in a folder</small>_

- After completion of the repair process, the software displays the repaired Excel file and its recoverable data in a preview window.

![preview of Excel file](https://www.stellarinfo.com/support/kb/images/Preview-of-excel-file.png)

                                                        _<small>Figure 7 - Preview of Excel File</small>_

- Click on **Save File** icon on **Home** menu to save the repaired file.
- In **Save File** dialog box, choose **Default location** or **Select New Folder** for saving the file.

![Select destination to save repaired excel file](https://www.stellarinfo.com/support/kb/images/select-destination-to-save-repaired-excel-file.jpg)

                                               _<small>Figure 8 - Select Destination to Save Repaired Excel File</small>_

- Click **OK** to proceed with the saving process.

The repaired file gets saved at the preferred location.

**_Note:_** _To recover the Engineering formulae, include ‘Analysis ToolPak’ Add-in._

 **Steps to Repair PowerPoint – ppt/pptx/pptm file**

- Click **Browse** to select the corrupt PowerPoint file. Alternately, click on **Search** to search for the file, if the location is not known.

![Select powerpoint presentation](https://www.stellarinfo.com/public/image/catalog/screenshot/powerpoint-repair/1-Stellar-Repair-for-Power-Point-Select-Corrupt-PPT-file.jpg)

                                    _<small>Figure 9 - Illustrates Selection of Single PowerPoint Presentation</small>_

- Once the corrupt PowerPoint file is selected, click **Scan** for scanning and repairing the file.
- A preview of scanned file gets displayed. Verify the file contents from the preview window.
- Click **Save** on **Home** menu to save the repaired PPT file.
- From the **Save File** dialog box, click **Default location** or **Other location** under **Save As** for saving the file.

![Save ppt](https://stellarinfo.com/support/kb/images/Select-location-to-save-ppt.png)

                                                    _<small>Figure 10 - Select Location to Save PPT File</small>_

- Click on the **OK** button and the repaired file is saved at preferred location.

**Steps to Repair PDF file**

- From the Stellar Repair for PDF main interface window, click **Add File** to select a single or multiple PDF files you want to repair.

![Adding corrupt pdf files](https://www.stellarinfo.com/screenshots/pdf-repair/1-Stellar-Phoenix-Repair-for-PDF-main-screen.jpg)

                                            _<small>Figure 11 - Illustrates adding of corrupt PDF Files</small>_

- A screen with recently added PDF file is displayed. Select the file and click **Repair** to start repairing it.

![Repair selected file](https://www.stellarinfo.com/screenshots/pdf-repair/2-Stellar-Phoenix-Repair-for-PDF-add-file.jpg)

                                                _<small>Figure 12 - Repair the Selected PDF File</small>_

- A screen showing the progress of the repair process appears.
- When the ‘Repair Complete’ window pops-up, click **OK**.
- Preview the repaired PDF file.
- Click the **Save Repaired Files** button to save the repaired file.

![save repaired file](https://www.stellarinfo.com/screenshots/pdf-repair/5-Stellar-Phoenix-Repair-for-PDF-preview.jpg)

                                                  _<small>Figure 13 - Save Repaired File</small>_

- In **Browse for Folder** dialog box, select a folder for saving the file.
- From the **Saving Complete** dialog box, click the hyperlink to the folder containing the repaired PDF file.

![saving complete Window](https://www.stellarinfo.com/screenshots/pdf-repair/7-Stellar-Phoenix-Repair-for-PDF-saved.jpg)

                                                      _<small>Figure 14 - Saving Complete Window</small>_

- Click **OK**.


## [Error Solved] Excel file is not in recognizable format

**Summary:** Microsoft’s Excel is one of the most widely used spreadsheet tools, however, it isn’t entirely free of errors. There are in fact quite a large number of problems that can crop up in this user-friendly application which can put all work to halt. One such error occurs when Excel does not recognize the file format of .xls or .xlsx file and the error message says “Excel file is not in recognizable format” error. Let us explore this annoying error in detail.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://cloud.stellarinfo.com/[StellarRepairforExcel-B.exe](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) "Free Download for Windows")

![Excel file is not in a recognizable format](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Excel-file-is-not-in-a-recognizable-format.jpg)

Figure: Error message

From a small shop to the global industry giants, everyone relies on Microsoft Excel to complete their work. Quite a few businesses not only use Excel for their inventory tracking purposes but also to manage task lists and timesheets for their employees and project management charts. With high programming proficiency, one can create macros in excel which help in automating a lot of things. You can create quite a few variations, such as pie charts, bar charts, line graphs, area charts, and many more to showcase the data both in a tabular column as well as in a pictorial representation.

While Excel enjoys wild popularity, thanks to its powerful design and features, it doesn’t mean that Excel is all free of errors. There are actually repetition a few errors that one can encounter. One you might have come across is the error stating “Excel file is not in a recognizable format”.

## **What is this error all about?**

The “Excel file in unrecognizable format error” occurs when the Excel file you are trying to load is corrupted. Microsoft has ensured that the workbook will be recoverable when the file is imported into excel but there are times when the automatic recovery does not happen. That’s where the challenge really lies. In such cases, getting to the root of the issue becomes necessary to be able to solve it.

## **Reasons behind the error**

1. One of the main reasons for the error is that the file must have got corrupted while being transferred from one machine to another.
2. Another reason can be that the latest service pack might not be in use on your system.
3. There could be MS Excel version change.
4. Corruption of the file due to virus infection, extremely large databases, or multiple locks on the file at the same time can also trigger this error.

If you have ever faced this error, you do not need to panic. We have a couple of solutions listed for you when you face the Excel file in an unrecognizable format error.

## **How do you go about fixing this?**

### **<u>Solution 1:</u> Use MOC.exe file to convert the workbook and then open it in Excel:**

1. Right-click on .XLS (you can use any .XLS files in your system).
2. A new dialogue will appear. Here, click on “Choose another app” to select it.

![Choose Another App](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Open-with.jpg)

Figure: choose another app

3. You will now be presented with a number of applications which the OS thinks the file format will be compatible with.
4. You do not have to choose any of the prepopulated apps from the list.

![Look for another app](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Look-for-another-app-on-this-PC.jpg)

Figure: Look for another app

5. Navigate using the **Look for another app on this PC**  to the path “C:\\Program Files\\Microsoft Office\\OfficeVersion”
6. You will see a file name MOC.exe
7. Choose that and complete your export.
8. Try opening the workbook in Excel and the error should now be resolved.

### **<u>Solution 2:</u> Opening the file from within the Excel:**

1. Open a new Excel workbook.
2. Press “Alt + F” or alternatively, go to the menu.
3. Once you are in the menu, go to **Options**.
4. You will be able to see a number of tabs on the left side of the options.
5. Under the ‘**Formulas**’ tab, ensure that the calculation is in Manual mode – this setting is in the automatic mode, by default.

![Manual option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Formula-option.jpg)

Figure: Manual option

6. Click **OK** and save the changes to the workbook.
7. Now, browse for the file which was corrupted.
8. Click on the file and then select the option “Open and Repair”. You will find it in the drop down Menu.

![Open and Repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Open-and-Repair.jpg)

Figure: Open and Repair

9. Once the file has been imported, click on “Repair” to recover the data from the selected workbook.

![Repair Option -Excel File](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/10/Repair-Excel-File.jpg)

Figure: Repair option

### **<u>Solution 3:</u> Use automated Excel repair software**

If none of the above mentioned manual methods works to eliminate the ‘Excel file in unrecognizable format’ error, it means your Excel file has been severely corrupted and needs professional assistance. In such a scenario, quickly download reliable and competent software [**Stellar Repair for Excel**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). Backed by powerful scanning and repair algorithms, this product guarantees up to 100% Excel file repair regardless of the amount of damage in it.

1. [**Download**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), install and launch Stellar Repair for Excel.
2. Allow the software to scan the corrupted Excel file.
3. All recoverable data will be listed in a tree-view list. You can select and preview any item from here.
4. Select and recover individual or entire data from the file and save as a new Excel.

This method is currently the easiest and most convenient to resolve miscellaneous Excel errors.

## **Wrapping it up**

Excel is one of the most powerful tools which can easily reduce your workload by more than 75% if used in a proper way. However, if you face complex errors like “Excel file is not in recognizable format”, you can use the methods mentioned above to get rid of it and resume your working in MS Excel. Remember, if the manual solutions don’t work, you can always rely on a proficient software like Stellar Repair for Excel to complete the job with finesse.


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
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-motorola-edge-40-neo-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Motorola Edge 40 Neo?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-tecno-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Tecno Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-videos-from-iphone-15-plus-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Videos from iPhone 15 Plus Without Backup? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-nokia-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Nokia</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-se-2022-without-losing-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone SE (2022) without Losing Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-8-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 8? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-excel-file-couldnt-open-in-protected-view-by-stellar-guide/"><u>How to Fix Excel File Couldnt Open in Protected View</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-xiaomi-14-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Xiaomi 14</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-contacts-on-iphone-6-plus-4-methods-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore Contacts on iPhone 6 Plus (4 Methods) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-realme-11x-5g-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Realme 11X 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-oppo-a58-4g-get-deleted-phone-number-back-with-ease-and-safety-by-fonelab-android-recover-contacts/"><u>How to Oppo A58 4G Get Deleted Phone Number Back with Ease and Safety</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-htc-u23-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from HTC U23?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-8-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone 8 Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-s18e-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from S18e</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-motorola-edge-40-neo-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Motorola Edge 40 Neo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-huawei-p60-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Huawei P60</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-honor-100-pro-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Honor 100 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-12-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone 12 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-12-mini-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone 12 mini Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-apple-iphone-15-plus-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Apple iPhone 15 Plus using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-poco-f5-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Poco F5 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-contacts-from-motorola-moto-g34-5g-by-fonelab-android-recover-contacts/"><u>How to get back lost contacts from Motorola Moto G34 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-c12-pro-by-fonelab-android-recover-video/"><u>How to recover old videos from your C12 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-galaxy-a25-5g-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Galaxy A25 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-dead-iphone-6s-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to recover data from dead iPhone 6s Plus | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-xr-without-apple-password-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone XR Without Apple Password? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-honor-magic-6-pro-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Honor Magic 6 Pro using Video Repair Utility?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-11-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 11 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-can-i-screen-mirroring-apple-iphone-se-2022-to-tvlaptop-drfone-by-drfone-ios/"><u>How Can I Screen Mirroring Apple iPhone SE (2022) to TV/Laptop? | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-detect-and-stop-mspy-from-spying-on-your-oneplus-ace-3-drfone-by-drfone-virtual-android/"><u>In 2024, How to Detect and Stop mSpy from Spying on Your OnePlus Ace 3 | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-transfer-data-from-apple-iphone-6s-to-others-ios-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From Apple iPhone 6s To Others ios devices? | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-does-airplane-mode-turn-off-gps-location-on-oppo-reno-10-proplus-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Does Airplane Mode Turn off GPS Location On Oppo Reno 10 Pro+ 5G? | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/guide-to-mirror-your-oppo-f23-5g-to-other-android-devices-drfone-by-drfone-android/"><u>Guide to Mirror Your Oppo F23 5G to Other Android devices | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-how-to-change-your-iphone-6-plus-apple-id-on-macbook-by-drfone-ios/"><u>In 2024, How To Change Your iPhone 6 Plus Apple ID on MacBook</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/list-of-pokemon-go-joysticks-on-nokia-xr21-drfone-by-drfone-virtual-android/"><u>List of Pokémon Go Joysticks On Nokia XR21 | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/best-fixes-for-vivo-y100a-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>Best Fixes For Vivo Y100A Hard Reset | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/how-to-remove-or-bypass-knox-enrollment-service-on-samsung-galaxy-a15-4g-by-drfone-android/"><u>How To Remove or Bypass Knox Enrollment Service On Samsung Galaxy A15 4G</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-pause-life360-location-sharing-for-poco-m6-pro-5g-drfone-by-drfone-virtual-android/"><u>In 2024, How To Pause Life360 Location Sharing For Poco M6 Pro 5G | Dr.fone</u></a></li>
<li><a href="https://android-location-track.techidaily.com/top-5-tracking-apps-to-track-oppo-k11x-without-them-knowing-drfone-by-drfone-virtual-android/"><u>Top 5 Tracking Apps to Track Oppo K11x without Them Knowing | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/how-to-reset-a-locked-itel-p40plus-phone-by-drfone-android/"><u>How to Reset a Locked Itel P40+ Phone</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-complete-tutorial-to-use-vpna-to-fake-gps-location-on-honor-x50-drfone-by-drfone-virtual-android/"><u>In 2024, Complete Tutorial to Use VPNa to Fake GPS Location On Honor X50 | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-fix-pokemon-go-route-not-working-on-vivo-y02t-drfone-by-drfone-virtual-android/"><u>How to Fix Pokemon Go Route Not Working On Vivo Y02T? | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/the-10-best-tools-to-bypass-icloud-activation-lock-from-apple-iphone-6s-you-should-try-out-by-drfone-ios/"><u>The 10 Best Tools to Bypass iCloud Activation Lock From Apple iPhone 6s You Should Try Out</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/additional-tips-about-sinnoh-stone-for-honor-x50i-drfone-by-drfone-virtual-android/"><u>Additional Tips About Sinnoh Stone For Honor X50i | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-how-to-remove-an-airtag-from-your-apple-id-account-from-iphone-13-by-drfone-ios/"><u>In 2024, How to Remove an AirTag from Your Apple ID Account From iPhone 13?</u></a></li>
<li><a href="https://unlock-android.techidaily.com/7-ways-to-unlock-a-locked-xiaomi-redmi-note-12r-phone-by-drfone-android/"><u>7 Ways to Unlock a Locked Xiaomi Redmi Note 12R Phone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/forgot-iphone-passcode-again-unlock-apple-iphone-se-2022-without-passcode-now-drfone-by-drfone-ios/"><u>Forgot iPhone Passcode Again? Unlock Apple iPhone SE (2022) Without Passcode Now | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/the-way-to-get-back-lost-call-history-from-infinix-note-30-pro-by-fonelab-android-recover-call-logs/"><u>The way to get back lost call history from Infinix Note 30 Pro</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-sharefake-gps-on-uber-for-xiaomi-13-ultra-drfone-by-drfone-virtual-android/"><u>How to share/fake gps on Uber for Xiaomi 13 Ultra | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-still-using-pattern-locks-with-itel-s23-tips-tricks-and-helpful-advice-by-drfone-android/"><u>In 2024, Still Using Pattern Locks with Itel S23? Tips, Tricks and Helpful Advice</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-fix-unfortunately-contacts-has-stopped-error-on-tecno-pova-5-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Fix Unfortunately, Contacts Has Stopped Error on Tecno Pova 5 | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-unlocking-iphone-se-lock-screen-3-foolproof-methods-that-actually-work-by-drfone-ios/"><u>In 2024, Unlocking iPhone SE Lock Screen 3 Foolproof Methods that Actually Work</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-call-history-from-motorola-moto-g73-5g-by-fonelab-android-recover-call-logs/"><u>Easy steps to recover deleted call history from Motorola Moto G73 5G</u></a></li>
</ul></div>



---
title: How to Fix Excel 2013 Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.534Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2013 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair .xlsm,repair .xls,repair .xltx,repair excel 2023,repair excel 2003
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





---
title: How to fix runtime error 424 object required error in Excel 2013 | Stellar
date: 2024-07-17T09:59:55.686Z
updated: 2024-07-18T09:59:55.686Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2013
excerpt: This article describes How to fix runtime error 424 object required error in Excel 2013
keywords: repair .xlsx files,repair corrupt .csv files,repair excel 2010,repair .xltm files,repair damaged .xls,repair damaged excel file,repair .xltx,repair .xlsm files,repair damaged .xlsm files,repair .xlsm,repair damaged .csv,repair .xltm
thumbnail: https://thmb.techidaily.com/07cfabd2fe9acb782e30cca8205dc0f557a2c3371dbf02532bc0633c00063d56.jpg
---

## How to fix runtime error 424 object required error in Excel

The Runtime error 424: Object required occurs when Excel is not able to recognize an object that you are referring to in a VBA code. The object can be a workbook, worksheet, range, variable, class, macro, etc. Some users have also reported that this error occurred when they tried to copy the values of the cells from one workbook to another.

Let’s understand the error through a small scenario. Suppose, I want to check the last field row in a table in a spreadsheet named “First” using the VBA code. To do this, I have added a command button and double-clicked on it and entered the below code in the backend:

Private Sub CommandButton2\_Click()

Dim LRow As Integer

LRow = Worksheets("First").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox ("Last Row " & LRow)

End Sub

![Code On Command Button](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/code-on-commandbutton.jpg)

In this code, _Worksheets("First")_ is a data object. If I mistakenly delete this data object and insert any random name (for example - kanada), then it will not be recognized by Excel. When I run this code, I will get the “Run-time error 424”.

![Runtime Error with scenario](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/run-time-error-424-with-scenario.jpg)

## **Causes of Runtime Error 424 in Excel**

The Runtime error 424: Object required can occur due to the following reasons:

- Incorrect name of the object you are trying to refer to in a code.
- You have provided an invalid qualifier to an object.
- You have not used the Set statement while assigning an object reference.
- The object is corrupted.
- Missing objects in a workbook.
- Objects you are trying to call in a code are mistakenly deleted or unavailable.
- You have used an incorrect syntax for object declaration.
- You are trying to perform an invalid action on an object in a code.
- Workbook is corrupted.

##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://aligracehair.sjv.io/c/5597632/2087264/19272" target="_top" id="2087264"><img src="//a.impactradius-go.com/display-ad/19272-2087264" border="0" alt="" width="336" height="280"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2087264/19272" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

<!-- affiliate ads begin -->
<a href="https://funwhole.sjv.io/c/5597632/1702887/17189" target="_top" id="1702887"><img src="//a.impactradius-go.com/display-ad/17189-1702887" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1702887/17189" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873305/18544" target="_top" id="1873305"><img src="//a.impactradius-go.com/display-ad/18544-1873305" border="0" alt="" width="1080" height="1350"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873305/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### 5\. Repair your Workbook

Sometimes, the ‘Object required’ error can occur if your Excel file is damaged or corrupted. In such a case, you can try repairing the file using Microsoft’s in-built utility - Open and Repair. To use this utility, follow these steps:

- In Excel, go to **File > Open > Browse**.
- In the Open dialog box, click on the corrupted Excel file.
- Click the arrow next to the Open button and select **Open and Repair** from the dropdown.
- Select **Repair** to recover as much data from the file as possible.

If the Open and Repair utility fails or stops working, then you can try a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can repair severely corrupted Excel files **(.xls, .xlsx, .xltm, .xltx, and .xlsm)**. It helps recover all the file components, including images, charts, tables, pivot tables, cell comments, chart sheets, formulas, etc., without impacting the original structure.

## **Conclusion**

The Runtime error 424 usually occurs when there is an issue with the objects in your VBA code. In this article, we have covered some effective methods to resolve the “object required” error in Excel. If the error occurs due to corruption in Excel file, then you can repair the corrupt file using Stellar Repair for Excel. It is a reliable tool that can repair severely corrupted Excel file without changing its actual formatting. You can download the free trial version of the software to evaluate its functionality.


<!-- affiliate ads begin -->
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033101&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Elite.png" border="0"></a>
<!-- affiliate ads end -->
## How to Repair Multiple Excel Files by Using Stellar

With Stellar Repair for Excel, it is quite easy and simple to repair multiple MS Excel (XLS and XLSX) files that are damaged. This is because the software has a self-explanatory interface and hence is a Do-it-yourself software. Nonetheless, when using this software to repair multiple Excel files, you would have to add all of the files into the software by following a few pre-defined steps. Follow the steps mentioned below:

- Launch **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** software.
- Under Home menu, click Select file

<!-- affiliate ads begin -->
<a href="https://sentrypc.7eer.net/c/5597632/398457/3022" target="_top" id="398457"><img src="//a.impactradius-go.com/display-ad/3022-398457" border="0" alt="www.sentrypc.com" width="980" height="120"/></a><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/398457/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Select file option](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/select-file-option.jpg)

- Click **Browse** and select corrupt Excel files. Select the checkbox to repair multiple files.

![Search file](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/22-search-file.png)

- Click Repair
- The software provides the preview facility. You can check the it on left pane.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851655&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
![Preview of file](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/4-preview.png)

- Save the repired filr ether **Default location** or **Select New Folder** radio button.

<!-- affiliate ads begin -->
<a href="https://godlikehost.sjv.io/c/5597632/1920047/21774" target="_top" id="1920047"><img src="//a.impactradius-go.com/display-ad/21774-1920047" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1920047/21774" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![select destination](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/6-save-file.jpg)

Stellar Repair for Excel Stellar Repair for Excel is the best choice for repairing corrupt or damaged Excel (.XLS/.XLSX) files. This Excel recovery software restores everything from corrupt file to a new blank Excel file.

[Learn More ![red arrow](https://www.stellarinfo.com/image/catalog/blacktheme/data-recovery-standard/red-arrow.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)


<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873313/18544" target="_top" id="1873313"><img src="//a.impactradius-go.com/display-ad/18544-1873313" border="0" alt="" width="1080" height="1263"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873313/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## How to fix Pivot Table Field Name is not Valid error in Excel?

The Pivot Table field name is not valid error can occur while creating, modifying, or refreshing data fields in the pivot table. It can also appear when using VBA code to modify the pivot table. It usually occurs when there is an issue with the field name in a code or if there is a hidden or empty column in the pivot table. However, there could be many other reasons behind this error.

## Why the "Pivot Table Field Name is not Valid" Error Occurs?

You can get the "Pivot Table field name not valid" error in Excel due to several reasons. Some possible causes are:

- Excel file is corrupted
- Damaged fields in the pivot table
- Pivot table is corrupted/damaged
- Hidden columns in the pivot table
- Macro (referring to the pivot table) is corrupted
- Preserve formatting option is enabled
- Missing or incorrect fields in the VBA code
- Issue with workbook.RefreshAll method syntax (if using)
- Pivot Table contains empty columns
- Header values or header column is missing in the Pivot Table
- Pivot table is created without headers
- Columns/rows are deleted from the Pivot Table

## Methods to Fix Pivot Table Field Name is not Valid Error in Excel

You can get this error if you have selected the complete data sheet and then trying to create the Pivot Table. Make sure you choose only the data fields that you want to insert in the Pivot Table. If this is not the case, then follow the troubleshooting methods mentioned below.

### **Method 1: Check the Header Value in the Pivot Table**

The "Pivot table field name is not valid" error can occur if you have not set up the pivot table correctly. All the columns having data in them should have header and header values. A pivot table without a header value can create issues. You can check the header and its value from the Formula bar. Change the header if the header value is too lengthy or if it contains special characters.

<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1821134/17882" target="_top" id="1821134"><img src="//a.impactradius-go.com/display-ad/17882-1821134" border="0" alt="" width="320" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1821134/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Adding reference for the document with details.](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/headers-value-in-formula-bars.jpg)

<!-- affiliate ads begin -->
<a href="https://aligracehair.sjv.io/c/5597632/2087267/19272" target="_top" id="2087267"><img src="//a.impactradius-go.com/display-ad/19272-2087267" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2087267/19272" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 2: Check and Change the Data Range in the Pivot Table**

The "Pivot Table field name is not valid" can occur while modifying a field in Pivot Table. It usually occurs if you're trying to add or modify the field by selecting an incorrect data range in the **Create PivotTable** dialog box. The **"Create PivotTable**" feature helps define how data would be displayed within the pivot table.

Let's take a scenario to understand this. Open the Excel file with PivotTable. Click on the fields (you want to add), go to the **Insert** option, and click **PivotTable.**  

![Inserting a Pivot Table from selection](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-insert-and-then-pivot-option.jpg)

If you select an incorrect range, i.e. A1:E18, instead of correct range - "Expenses**!$A$3:Expenses!$A$4**," you will immediately get the error message.

![Selecting a table range with values for report](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/select-table-with-correct-range.jpg)

So, type the correct range under the Select a table or range option and click **OK**.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3727260&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
### **Method 3: Unhide Excel Columns/Rows**

The error can also occur if some columns/rows of the Pivot Table's data source are hidden. When you try to add a hidden column as a field in the PivotTable, the Excel application will fail to read the data of the hidden column. You can check and unhide the Excel columns by following these steps:

- Open the Excel file.
- Locate the hidden column number.
- Move your cursor on the hidden column number and right-click on the space between the columns. Click **Unhide**.  

    ![unhiding the rows in Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-unhide-option.jpg)

<!-- affiliate ads begin -->
<a href="https://atezr.pxf.io/c/5597632/2018605/18496" target="_top" id="2018605"><img src="//a.impactradius-go.com/display-ad/18496-2018605" border="0" alt="" width="798" height="807"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2018605/18496" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 4: Check and Delete Empty Excel Columns**

Sometimes, you can get the "Pivot Table field name is not valid" error if you are trying to use an empty column as a field in your Pivot Table. Check the columns with no values in all cells. If found, then delete the empty columns. This method is ideal for small-size Excel files. However, for large-sized files, it is a time-consuming process.

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=27889512&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/728__90.jpg" border="0"></a>
<!-- affiliate ads end -->
### **Method 5: Unmerge the Column Header (If Merged)**

The "Pivot Table field name is not valid" error can also occur due to merged column headers. The pivot table references headers to identify the data inside the rows or columns. The merged headers can sometimes create data inconsistencies. You can try unmerging the column headers to fix the issue. Follow these steps:

- In the Excel file, go to the **Home**
- Click the **Merge & Center** option and select **Unmerge Cells** from the dropdown.  

    ![unmerging cells from home tab in Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-home-select-unmergecells.jpg)

### **Method 6: Disable the Background Refresh Option**

If the "background refresh" option in the Excel file is enabled, it may also create issues with Pivot Table. The Excel updates all the pivot tables in the background even after a small change if the background refresh option is enabled. This may create issues if the Excel file is large with too many tables. You can try turning off the "background refresh" option in the Excel file to troubleshoot the issue. Here is how to do so:

- In the Excel file, go to the **Data** tab and then click **Connections**.  

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087484/7443" target="_top" id="2087484"><img src="//a.impactradius-go.com/display-ad/7443-2087484" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087484/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
    ![Adding connections from the data](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/go-to-data-then-click-connections.jpg)

- In the **Workbook Connections**dialog box, click on the **'Add'** dropdown to add the workbook (in which you need to modify the refresh settings).  

<!-- affiliate ads begin -->
<a href="https://bluettieu.pxf.io/c/5597632/2042323/17091" target="_top" id="2042323"><img src="//a.impactradius-go.com/display-ad/17091-2042323" border="0" alt="BLUETTI NEW LAUNCH AC180T" width="3840" height="1600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2042323/17091" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
    ![Add the option for the Workbook connections.](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-add-from-drop-down.jpg)

- Once you have chosen the Excel file, click **Properties.**  

<!-- affiliate ads begin -->
<a href="https://ancheer.sjv.io/c/5597632/1657301/17326" target="_top" id="1657301"><img src="//a.impactradius-go.com/display-ad/17326-1657301" border="0" alt="" width="1920" height="933"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1657301/17326" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
    ![Selecting Properties for the Workbook connections.](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-properties-on-workbook-connections.jpg)

- In the **Connection Properties** window, unselect the **"Enable background refresh"**option, select the "**Refresh data when opening the file**", and click **OK.  

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/2016067/18544" target="_top" id="2016067"><img src="//a.impactradius-go.com/display-ad/18544-2016067" border="0" alt="" width="1020" height="380"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2016067/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
    ![Enabling the connection properties by enabling and refreshing data](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/select-background-refresh-and-refresh-data-option.jpg)

    **

### **Method 7: Check the VBA Code**

The error can also occur when working with PivotTable using VBA code in Excel. Some Excel users reported this error on forums as **run-time error 1004: The PivotTable field name is not valid**. This error usually occurs when there are issues in the VBA code, affecting the PivotTable data source or field references. You can check field names referring to PivotTable or Workbook.RefreshAll function syntax and other errors in the code.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068425/7443" target="_top" id="2068425"><img src="//a.impactradius-go.com/display-ad/7443-2068425" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068425/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 8: Repair your Excel File**

One of the reasons behind the "Pivot Table field name is not valid" error is corruption in the Excel file, containing the Pivot Table. You can repair your Excel file using Microsoft built-in utility - Open and Repair. Here's how to use this utility:

- In Excel, navigate to **File > Open.**
- Click **Browse** to choose the affected workbook.
- The **Open** dialog box will appear. Click on the corrupted file.
- Click the arrow next to the **Open**button and then select **Open and Repair**.
- You will see a dialog box with three buttons - **Repair, Extract Data,** and **Cancel**.  

    ![Repairing the corrupt workbook from Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivotimage/click-repair-option.jpg)

- Click on the **Repair** button to recover as much of the data as possible.
- After repair, a message is displayed. Click **Close**.

###  **Method 9: Use a Professional Excel Repair Tool**

If the Excel file is heavily damaged or corrupted, then the "[Open and Repair" utility may not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) or provide the intended results. In such a case, you can opt for a professional Excel repair tool. **Stellar Repair for Excel** is an advanced Excel file repair tool, which is highly recommended by experts. It can repair severely corrupted Excel files and restore all the data from corrupt file, including pivot tables. This tool comes with a user-friendly interface that even a non-technical user can use. You can try the software's demo version to check how it works. The software is fully compatible with all Excel versions, including Excel 2019.

## **Conclusion**

The Excel error "Pivot Table field name is not valid" can occur due to hidden or merged column/row headers, empty columns/rows, corrupted pivot table, and various other reasons. You can try the methods mentioned above to fix the error. If this error has occurred due to corruption in the Excel file, then you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) - an advanced tool to repair corrupted pivot table, macros, fields, or other elements in an Excel file. It is compatible with all Windows editions, including the latest Windows 11. It can help fix the error if the data source or Pivot table configuration is affected by corruption.


<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075471/7443" target="_top" id="2075471"><img src="//a.impactradius-go.com/display-ad/7443-2075471" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075471/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Repair Office 2016 Files (Word, Excel and PowerPoint)on Windows

If you frequently work with Microsoft Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files, then issues like file inaccessibility or corruption won’t be new to you.

Let’s discuss some common scenarios which may lead to corrupt MS Office 2016 files:  

## Scenarios behind Microsoft Office Files Corruption

<!-- affiliate ads begin -->
<a href="https://arkmc.pxf.io/c/5597632/427477/5172" target="_top" id="427477"><img src="//a.impactradius-go.com/display-ad/5172-427477" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://arkmc.pxf.io/i/5597632/427477/5172" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Scenario 1 – Disruption during Data Migration

You decide to move Office files from your hard drive to other removable media. However, when you try to access the data within the files post-migration, you may find Word, Excel, and PowerPoint files showing gibberish characters. Due to a power surge, sudden system shutdown, and internal mechanical failure, the files may have turned corrupt.

![](https://www.stellarinfo.com/image/catalog/article/word/Word-displaying-gibberish-characters.png)

Figure 1- Microsoft Word file showing garbage characters

### Scenario 2 – Office Files and Registry Entries Become Infected

When you open or use the Microsoft Office application, it crashes as soon as it opens. You assume that an add-in was causing the problem and restart the Office application without add-ins loaded, but the application still crashes. This may happen because of a virus infecting the Office files and registry values, thus leading to corrupt or damaged Office files.

### **Scenario 3 – Inaccessible or Lost Data**

Suppose all your Office files are stored on a USB device, and you unplugged the device while it was still open in Windows. Now, when you attempt to open a Word or an Excel file, all the data is gone. Unsafe removal of USB or any other external storage device may corrupt the data inside your Office files or turn the file inaccessible.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4940312&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/333ac5d90817d69113471fbb6e531bee/sps-partnership-728x90eng.png" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://tokenmetrics.sjv.io/c/5597632/1864921/20702" target="_top" id="1864921"><img src="//a.impactradius-go.com/display-ad/20702-1864921" border="0" alt="" width="1251" height="1042"/></a>
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/image/catalog/article/word/ControlPanel-Uninstall.png)

- Right-click on the Office application you want to repair, and then click **Change**

<!-- affiliate ads begin -->
<a href="https://proteahair.pxf.io/c/5597632/1983634/23621" target="_top" id="1983634"><img src="//a.impactradius-go.com/display-ad/23621-1983634" border="0" alt="" width="320" height="100"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1983634/23621" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/image/catalog/article/word/Change-Office-application.jpg)

- Under **Change your installation of Microsoft Office Professional Plus 2016,** choose Repair and then click **Continue.**

![](https://www.stellarinfo.com/image/catalog/article/word/Repair-Office-Application.jpg)

**For Windows 10**

- Right-click the Start button, and type in **Apps & Features** (For Windows 10)

![](https://www.stellarinfo.com/image/catalog/article/word/Apps-and-features.png)

**NOTE:** This step will work for Windows 10/8/8.1/7 and Vista

- Click **Programs** from the window that opens, click on the MS Office product you want to repair, and then click on **Modify**

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/2072819/17059" target="_top" id="2072819"><img src="//a.impactradius-go.com/display-ad/17059-2072819" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072819/17059" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/image/catalog/article/word/Modify-Office-application.jpg)

**Note:** Following the step will repair the entire Microsoft Office suite even if it contains only one application you want to repair such as an Excel or PowerPoint file. But, in case you have a standalone app installed, try to locate that application by name.

- Under **Change your installation of Microsoft Office Professional Plus 2016,** choose Repair, and then click **Continue** to initiate the repair process.

<!-- affiliate ads begin -->
<a href="https://caperobbin.sjv.io/c/5597632/2006123/18460" target="_top" id="2006123"><img src="//a.impactradius-go.com/display-ad/18460-2006123" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2006123/18460" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/image/catalog/article/word/Repair-Office-Application.jpg)

- Once the repair process completes, you'll be prompted to restart your PC. Click **Yes**

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BGeneral%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/general-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BBusiness%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/business-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
## Conclusion

This post outlined possible scenarios and their causes that may lead to corruption in MS Office 2016 files. It also emphasized how the inbuilt methods such as Open and Repair, and Repair Office Installation help to resolve the corruption issues. But these are not competent enough to resolve all the errors. With Stellar Toolkit for File Repair, you can resolve all sorts of corruption issues and recover data of Office 2016 files – Excel, Word, PPT, and PDF – in their original state.




## \[Fixed\] Excel Found a Problem with One or more Formula

**Summary:** The error ‘Excel found a problem with one or more formula references in this worksheet’ may appear while saving the Excel workbook. It occurs when Excel found a problem with the formula used in the sheet. However, it may also occur when the Excel workbook gets damaged or corrupt. In this guide, we’ve explained the reasons that may lead to this Excel error and methods to resolve the error, by using various Excel options and a third-party Excel file repair software.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

If you are experiencing the ‘Excel found a problem with one or more formula references in this worksheet’ error message in the Excel workbook, it indicates that the [Excel file is corrupt](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) or partially damaged. However, it may also occur due to incorrect reference to a wrong cell or object linking, which is not working. The complete error message says,

_‘Excel found a problem with one or more formula references in this worksheet. Check that the cell references, range names, defined names, and links to other workbooks in your formulas are all correct.’_

![Excel found a problem with one or more formula references](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-problem-with-formula-reference.png)

In any case, resolving the error is critical as it doesn’t let you save the file and may result in loss of information from the Excel workbook.

<!-- affiliate ads begin -->
<a href="https://ukaidot.sjv.io/c/5597632/1793237/19578" target="_top" id="1793237"><img src="//a.impactradius-go.com/display-ad/19578-1793237" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793237/19578" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Reasons for Excel Formula References Error

A few reasons that may lead to such error are as follows,

- Wrong formula or reference cell
- Incorrect object linking or link embedding OLE
- Empty or no values in named or range cells
- Multiple Excel files (not common)

## Methods to Resolve ‘Excel Found a Problem with One or More Formula References in this Worksheet’ Error

Following are a few methods that you can follow to [fix Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that can’t be saved due to problems with one or more formula references in the worksheet.

<!-- affiliate ads begin -->
<a href="https://imp.i110150.net/c/5597632/924299/11305" target="_top" id="924299"><img src="//a.impactradius-go.com/display-ad/11305-924299" border="0" alt="" width="520" height="100"/></a>
<!-- affiliate ads end -->
### Method 1: Check Formulas

If the problem has occurred in a large Excel workbook with multiple sheets, it’s quite hard to pinpoint the problem cell. In such cases, you can use the Error Checking option that runs a scan and checks for a problem with formulas used in the worksheet.

To run Error Checking in the Excel sheet, follow these steps,

- Go to Formulas and click on the ‘Error Checking’ button

![Error Checking](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-error-checking-1024x431.png)

- This runs a scan on the sheet and displays the issues, if any. If no issue is found, it displays the following message,

_The error check is completed for the entire sheet._

In such a case, you can try saving the Excel file again. If the error message persists, proceed to the next method.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851691&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
### Method 2: Check Individual Sheet

The problem may also occur due to an issue with one of the sheets in the workbook. To find the faulty sheet and fix the problem, you can copy each sheet content in a new Excel file and then try to save the Excel file.

This will help you find the faulty sheet from the workbook that you can review. This method makes the entire [process of troubleshooting Excel formula](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) reference error quite easy and convenient.

In case the error is not fixed, you can back up the faulty sheet content and remove it from the workbook to save the Excel file.

<!-- affiliate ads begin -->
<a href="https://tinyland.pxf.io/c/5597632/1793214/19135" target="_top" id="1793214"><img src="//a.impactradius-go.com/display-ad/19135-1793214" border="0" alt="" width="900" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793214/19135" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Method 3: Check Links

When the Excel file contains external links with errors, MS Excel may display such error messages. To check and confirm if external links are causing the error, follow these steps,

- Navigate to _Data Tab > Queries & Connections > Edit Links_
- Check the links. If you find any faulty link, remove it and then save the sheet

### Method 4: Review Charts

You can review the charts to check if they are causing the formula reference error in Excel. It may take a while based on the size of the Excel file. Sometimes, it’s not practically possible to track down which Excel chart object is causing the error. Thus, you need to check specific locations, such as:

1. Check horizontal axis formula inside Select Data Source dialog box
2. Check Secondary Axis
3. Check linked Data Labels, Axis Labels, or Chart Title

### Method 5: Check Pivot Tables

To check Pivot Tables, follow these steps,

- Navigate to _PivotTable Tools > Analyze > Change Data Source > Change Data Source…_

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075482/7443" target="_top" id="2075482"><img src="//a.impactradius-go.com/display-ad/7443-2075482" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075482/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Edit links](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-Edit-Links-1024x84.png)

- Check if any of the formula used is problematic. Sometimes small typo, such as misplaced comma, can lead to such problems in Excel. Thus, check each formula thoroughly and correct the formulas wherever needed.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BGeneral%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/general-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### Method 6: Use Excel Repair Software

When none of the methods resolve the error, then you can rely on advanced [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. It’s a powerful tool that is recommended by several MVPs and IT administrators for resolving common Excel errors, such as ‘Excel found a problem with one or more formula references in this worksheet.’

![Stellar Repair for Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/public/image/catalog/screenshot/excel-repair/stellar-repair-for-excel-main-interface.png)

It repairs corrupt or damaged Excel (.xls/.xlsx) files, recovers Pivot tables, charts, etc., and save them in a new Excel worksheet. It helps Excel users, facing formula reference error, restore their Excel file without any risk of data loss, while preserving the sheet properties and formatting with 100% precision.

## Conclusion

Although the error ‘Excel found a problem with one or more formula references in this worksheet’ can be resolved by using various options in MS Excel, it may lead to a partial loss of information. Thus, you must perform these operations after taking a backup of the Excel worksheet. Also, if the MS Excel options fail to resolve the problem, you can use an Excel file repair software, such as Stellar Repair for Excel. The software helps fix Excel file corruption and restores the information and data from corrupt or damaged Excel files (.xls/.xlsx) to a new worksheet.




## Simple ways to Open Corrupt Excel file Without any Backup

**Summary:** The blog describes simple ways to open corrupt Excel file without any backup. It explains some manual workarounds that you can try to open the file. Also, it mentions about an Excel file repair tool that can quickly fix the corrupt file and recover data from it.

<!-- affiliate ads begin -->
<a href="https://mushroom-supplies.sjv.io/c/5597632/1692242/18134" target="_top" id="1692242"><img src="//a.impactradius-go.com/display-ad/18134-1692242" border="0" alt="" width="834" height="592"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1692242/18134" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Do you have an Excel file that does not open because of corruption issue? And every time you try to open it, an error message ‘the file is corrupt and cannot be opened’ pops-up?

<!-- affiliate ads begin -->
<a href="https://bluettius.sjv.io/c/5597632/2027209/17108" target="_top" id="2027209"><img src="//a.impactradius-go.com/display-ad/17108-2027209" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2027209/17108" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Excel file is corrupt and cannot be opened message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Excel-file-corruption-message-300x139.png)

Excel File Corruption Message

Also, you don’t have a healthy backup of the Excel file to restore the data? If so, you can try repairing the corrupt file by using a few simple yet effective manual workarounds mentioned below.

## **How to Open a Corrupt Excel File without Backup?**

Following are some manual methods that can help you open a corrupt Excel file:

### **Method 1: Repair Corrupt Excel File**

When attempting to open a corrupt file, Excel automatically starts ‘File Recovery’ mode to repair the file. But, if the recovery mode doesn’t start, try Microsoft Excel’s built-in ‘Open and Repair’ feature to manually repair the file.

To use this feature, perform the following steps:

**Step 1:** Open a **Blank workbook** in Excel, and then click **File > Open**.

**Step 2:** In the **Open** window, browse and select the corrupt file.

**Step 3:** Click the arrow that is beside the **Open** tab, and select **Open and Repair**.

![Open a blank workbook in Excel, navigate to File > Open, choose the corrupt file, and, in the Open window, click the arrow beside the Open tab, selecting Open and Repair for file recovery.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Open-and-repair.png)

Open and Repair Option

**Step 4:** Implement one of the following:

- Click the **Repair** button. (This is to recover as much data as possible.)
- Click the **Extract Data** button. (This is to recover values and formulas from the Excel file if the repair process fails to recover the entire data.)

![Initiate file recovery by selecting the Repair tab, and if necessary, retrieve values and formulas using the Extract Data tab in Excel.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/repair-excel-file-1-768x158.jpg)

Excel Built-in Repair Options

If using [Open and Repair does not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), move to the next method.

<!-- affiliate ads begin -->
<a href="https://25home.pxf.io/c/5597632/2090698/16836" target="_top" id="2090698"><img src="//a.impactradius-go.com/display-ad/16836-2090698" border="0" alt="" width="720" height="300"/></a>
<!-- affiliate ads end -->
### **Method 2: Disable the Protected View Feature**

Some Excel users have reported that turning off the ‘protected view’ feature in Excel helped them open the corrupt file. You can also try to disable this feature and open the file. To do so, follow these steps:

**Step 1:** Open a blank Excel file, click on **File** > **Options**.

**Step 2:** In the **Excel Options** window, select **Trust Center**, and then click **Trust Center Settings**.

<!-- affiliate ads begin -->
<a href="https://godlikehost.sjv.io/c/5597632/1920054/21774" target="_top" id="1920054"><img src="//a.impactradius-go.com/display-ad/21774-1920054" border="0" alt="" width="320" height="100"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1920054/21774" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![In the Trust Center tab, click on Trust Center Settings...](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Excel-trust-center-settings.png)

Excel Trust Center Settings

**Step 3:** Click **OK.**

Now check if you can open the corrupt file. If not, try implementing the next method.

### **Method 3: Look For Automatically Recovered Excel File**

If you have Excel’s AutoRecover feature enabled, you’ll have access to a copy of the **Excel file corrupted** or lost due to application crash, power outage, or accidental deletion.

**The ‘AutoRecover’** feature saves Excel worksheets at a temporary location after a certain time interval. It saves the worksheets automatically and is turned on by default to reduce the chance of data loss.

Check if you can **[recover corrupted Excel file](https://support.microsoft.com/en-us/office/repair-a-corrupted-workbook-153a45f4-6cab-44b1-93ca-801ddcd4ea53)** by following these steps:

**Step 1:** In Excel, open a **Blank workbook**.

**Step 2:** Go to **File** and click **Options**.

![Open a new Excel workbook, then access additional settings by navigating to File and selecting Options.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/08/Select-options-in-Excel-2013.jpg)

Figure 5 – Excel Options

**Step 3:** In the **Excel Options** dialog box, click **Save**, and then copy the ‘AutoRecover file location’.

![Copy the 'AutoRecover file location' for configuration or backup purposes.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Autorecover-excel-file.png)

Excel Options Window

**Step 4:** Open File Explorer window and paste the copied AutoRecover file location, and press **Enter**.

**Step 5:** A list of saved Excel files will be displayed. Choose the file you want to recover.


_**TIP:** Use Excel’s AutoBackup feature to reduce chances of data loss, by saving a previous version of your spreadsheet automatically._

## **Use an Excel File Repair Software**

If the above manual methods fail, repair the **corrupt Excel file** by using a third-party software, such as Stellar Repair for Excel**.** The software helps repair Excel (XLS and XLSX) files easily and effectively.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087389/7443" target="_top" id="2087389"><img src="//a.impactradius-go.com/display-ad/7443-2087389" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087389/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**Read this: [How to repair corrupt Excel file using Stellar Repair for Excel?](https://www.stellarinfo.com/support/kb/index.php/article/repair-corrupt-excel-file)**

Some key features of Excel Repair software are as follows:

- Fixes all errors in the MS Excel file.
- Repairs multiple damaged Excel files in a go.
- Recovers chart, chart sheet, table, cell comment, image, formula, and sort & filter.
- Preserves properties and cell formatting of Excel worksheets.
- Previews recoverable Excel file data before saving.
- Recovers all data components from the corrupt files and saves them in a new blank Excel file.
- Compatible with Excel 2019, 2016, 2013, 2010, 2007, and lower versions.

<!-- affiliate ads begin -->
<a href="https://unicoeye.pxf.io/c/5597632/2084396/18498" target="_top" id="2084396"><img src="//a.impactradius-go.com/display-ad/18498-2084396" border="0" alt="" width="1920" height="700"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084396/18498" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Conclusion

You can try the workarounds discussed in the blog to open a corrupt Excel file without a backup. Disabling the protected view feature can help you open the file. If the issue persists then try repairing the corrupted Excel file using the Open and Repair utility. Although, it may not be able to fix a severely corrupted workbook. In such a case you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can help you repair a corrupted Excel file with 100% integrity.


## Recover Corrupted Excel File 2007, 2010 | Easy Methods

There are several reasons that can cause Microsoft Excel workbooks to turn corrupt, such as virus attack, bad sectors on a drive on which Excel file is saved, system shutdown without properly closing the Excel application, etc.

Corruption in an Excel workbook can result in data loss or render the workbook inaccessible. Fortunately, Excel automatically starts recovery upon opening a corrupted Excel file. But, if it fails, you can manually repair the file or extract data from the corrupt file.

Quick Solution: Performing 2007, 2010 Excel repair or recovery process manually can be time-consuming. Also, manual workarounds to recover corrupt Excel workbook does not guarantee recovering the complete workbook data. Use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) software to repair single or multiple Excel (XLS/XLSX) files in 3 simple steps. The software also helps recover the Excel file, keeping the data intact.

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=4596923&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_468X60.png" border="0"></a>
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/image/catalog/article/Quick-Way-to-Fix-MS-Excel-2007---2010%20(1).jpg)

## **How to Fix** **Microsoft Excel 2010 & 2007 Files Corruption?**

Microsoft Excel comes with an inbuilt repair utility, called ‘Open and Repair’, that helps fix and recover corrupted Excel files.


### **Steps to Repair MS Excel 2010 Files Manually**

The detailed steps to open and repair Excel 2010 are as follows:

- Open Microsoft Excel 2010 and click **File** from the main menu.

![Microsoft Excel 2010 File Menu](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img1.JPG)

- Next, click **Open**.

<!-- affiliate ads begin -->
<h3 id="200610"><a href="https://sentrypc.7eer.net/c/5597632/200610/3022">Parental Control Software</a></h3>
<span class="text-ad-content">
	#1 Rated Parental Control Software.<br/>
	Monitor & Control all PC Activity!<br/>
		<cite style="color:green">sentrypc.com/parental-controls/</cite>
	</span><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/200610/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Select Open Option](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img2.JPG)

- Browse the corrupt Excel 2010 file on your computer and select it in the Open dialog.

![Browse Corrupt Excel 2010 File](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img3.JPG)

- Click the arrow next to the **Open** button and choose **Open and Repair**.

![Select Open and Repair](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img4.JPG)

- Click **Repair** when prompted to recover data to the maximum.
- If Excel fails to repair, click **Extract Data** to extract values and formulas in the corrupt file.
- Excel prompts to 'Convert to Values' or 'Recover Formulas'.
- Click **Yes** if it prompts the following error:

**_"The document file name caused a serious error the last time it was opened. Would you like to continue opening it?_**

- When Excel opens the last saved file, save it.

Once you’re able to access the last saved 2010 Excel file, try extracting the file contents.

<!-- affiliate ads begin -->
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033095&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Advanced-3YR.png" border="0"></a>
<!-- affiliate ads end -->
### **Save Excel 2010 File in HTML Format**

If you can open the Excel file, choose the HTML format to save it in filtered form. After that, close the Excel file as you have your data in the HTML file. The steps to save an Excel file in HTML format are as follows:

- Open Microsoft Excel 2010, click **Save As**, and then choose **Web Page** in the ‘Save as’ type drop-down list.
- Select the "Enable Entire Workbook” option, and then click the **Save** button.
- Close the Excel file and reopen your Microsoft Excel application. Browse the HTML file that you have saved.
- Click **File** from the main menu, and select **Save As** in the list.
- Type-in a different name, choose Microsoft Excel Workbook in the ‘Save as’ type drop-down menu, and then click the **Save** button.

With this, you would be able to access the data in the corrupt Excel file.

If the inbuilt tool fails to repair Excel 2010 file, a few methods can help you recover data from corrupted or lost workbook manually.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087407/7443" target="_top" id="2087407"><img src="//a.impactradius-go.com/display-ad/7443-2087407" border="0" alt="" width="600" height="500"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Steps to Repair Excel 2007 Files Manually**

Follow these steps to repair a corrupted 2007 Excel file by using the inbuilt Microsoft Excel repair tool:

- Open Microsoft Excel 2007, click the **Office button**, and then select **Open**.

<!-- affiliate ads begin -->
<a href="https://thefitville.pxf.io/c/5597632/1526796/15852" target="_top" id="1526796"><img src="//a.impactradius-go.com/display-ad/15852-1526796" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1526796/15852" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Open Microsoft Excel 2007 Main Menu](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img5.JPG)

- In the Open dialog box that pops-up, browse and select the corrupt Excel 2007 file. Click the arrow next to the **Open** button and choose **Open and Repair.**

**![Open and Repair Excel 2007 File](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img6.JPG)**

- Click **Repair** when prompted to recover as much data as you can from Excel 2007 file.

![Repair Excel 2007 File](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img7.JPG)

- If a repair fails, follow steps 1 till 3, and then click **Extract Data** to extract values and formulas from the corrupt file.
- In the window that appears, click **Convert to Values** or **Recover Formulas** to extract workbook data.

![Recover Excel 2007 File](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img8.JPG)

**_Note:_** _The built-in Microsoft Excel 2007 repair tool may fail to resolve Excel corruption issue. Also, choosing to extract data from the workbook involves data loss risk. Using a professional Excel repair tool, however, can resolve all types of_ [_Excel file corruption errors_](https://www.stellarinfo.com/support/kb/index.php/article/resolve-excel-file-corruption-errors) _and restore all its data._

<!-- affiliate ads begin -->
<a href="https://uperfect.sjv.io/c/5597632/1246754/15155" target="_top" id="1246754"><img src="//a.impactradius-go.com/display-ad/15155-1246754" border="0" alt="" width="600" height="600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1246754/15155" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Methods to Recover Data from Corrupt Excel 2010 & 2007 Files**

If the ‘Open and Repair’ feature fails in getting your Excel 2010, 2007 file repaired, you can try retrieving the file contents by following some manual methods. However, the methods may vary depending on whether you can open a workbook or not.

<!-- affiliate ads begin -->
<a href="https://homestyler.sjv.io/c/5597632/2044747/22993" target="_top" id="2044747"><img src="//a.impactradius-go.com/display-ad/22993-2044747" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2044747/22993" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 1 – Move Corrupt Excel File to another System**

Move the corrupt Excel file to any other computer and try opening it in MS Excel 2010/2007. Doing so, may help you resolves disk or network-related errors leading to Excel file corruption.

### **Method 2 – Revert Unsaved Excel File to its Last Saved Version**

If an Excel file turns corrupt while working on it but before saving any changes, try reverting it to its last saved version. To do so, perform the following:

- Open your Excel application, click the **Office button**, and then click **Open** from the menu.
- Browse the corrupt Excel file, click **Yes** when prompted to revert to its last saved version.

## **What if Nothing Works?**

If you fail to recover a corrupt Excel 2007/2010 file, perform Excel file recovery with [**Stellar Excel repair software**.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) The software is specially designed to help users fix their corrupted XLS/XLSX files quickly and easily without any technical assistance. It also helps restore all the file data to its original form.

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/1997817/17059" target="_top" id="1997817"><img src="//a.impactradius-go.com/display-ad/17059-1997817" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1997817/17059" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free download Stellar Repair for Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Points to Remember  

- Close all the MS Excel instances before using the software
- If the sheet you are repairing contains engineering formulas, please include ‘Analysis TooPak’ manually from Tools > Add-Ins

If you know the corrupt Excel 2007 or 2010 file location, click **Browse** to choose the file. Otherwise, click **Search**. Follow the below steps to recover data from corrupt Excel 2007/2010 file by using Stellar Excel repair tool:  

<!-- affiliate ads begin -->
<a href="https://getlyla.pxf.io/c/5597632/1455723/15391" target="_top" id="1455723"><img src="//a.impactradius-go.com/display-ad/15391-1455723" border="0" alt="" width="336" height="280"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1455723/15391" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Select Corrupt Excel File in Stellar Repair for Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img9.JPG)

- Click the **Repair** button to scan the file.

![Repair Excel File using Stellar Repair for Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img10.JPG)

- Once the scanning process is complete, the software shows a preview of recoverable Excel file items.

![Preview of Recoverable Excel File Items](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img11.JPG)

- To save the repaired file, click the **Save File** option on **File** menu.

![Choose Save File](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img12.JPG)

- In ‘Save File’ dialog box, choose to recover Excel 2007 & 2010 data to either the Default or New location. Click **OK**.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082541/7443" target="_top" id="2082541"><img src="//a.impactradius-go.com/display-ad/7443-2082541" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082541/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![File Saving Options](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img13.JPG)

The repaired Excel file gets saved at the specified location.

## **Preventive Measures to Avoid Losing Excel File Data**

The above-discussed methods might help salvage your data. But, it is recommended that you must take some preventive measures to avoid losing the data. One such important measure is backing up a copy of your workbook automatically. Doing so, will help you get back data in case the workbook is accidentally deleted or corrupted.

### **Steps to Create Backup Copy Automatically**

You can automatically create an Excel backup copy by following these steps:

- Click **Save As** from the main menu of your Excel application.
- **Browse** to the location where the corrupt Excel 2010/2007 file is saved.

![Browse the Excel File Location](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img14.JPG)

- In ‘Save As’ dialog box, click the arrow next to **Tools** button (given at the bottom left corner) and choose **General Options**.

![Choose General Options](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img15.JPG)

- In ‘General Options’ box, check **Always create backup** checkbox, and then click **OK**.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B300x600%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-300x600.gif" border="0"></a>
<!-- affiliate ads end -->
![Select Always Create Backup](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/img16.JPG)

With this, you instructed MS Excel to create a backup of every Excel file you create or open for work.

## **Conclusion**

This article outlined the typical reasons resulting in a corrupt Excel 2010 or 2007 file, such as virus infection, bad sectors on drive, etc. It explained how to fix a corrupted Excel file by using the inbuilt MS ‘Open and Repair’ tool. The article also discussed methods to recover Excel files in MS Office 2010 & 2007 when the Microsoft Excel repair tool fails. Further, it explained how using a professional repair tool such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) can come in handy when the manual methods to repair and recover Excel 2007 and 2010 file fails. But, keep in mind, a workbook may get corrupt again. And so, make sure to automatically backup your workbook to avoid losing its data.


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







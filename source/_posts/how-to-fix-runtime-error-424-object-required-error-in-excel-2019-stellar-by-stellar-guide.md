---
title: How to fix runtime error 424 object required error in Excel 2019 | Stellar
date: 2024-06-25T04:05:29.427Z
updated: 2024-06-26T04:05:29.427Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2019
excerpt: This article describes How to fix runtime error 424 object required error in Excel 2019
keywords: repair corrupt .xlb,repair damaged .xlb files,repair .xlb files,repair .xlsx files,repair corrupt .xlsm files,repair corrupt .xltm files,repair excel 2019,repair damaged .xlsm files,repair corrupt .xls,repair .csv files,repair excel 2000
thumbnail: https://thmb.techidaily.com/035705869a176d12c457c62dcd5ac8433382a242da2e6ee8d5c9aeccc24af52d.jpg
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

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

### 5\. Repair your Workbook

Sometimes, the ‘Object required’ error can occur if your Excel file is damaged or corrupted. In such a case, you can try repairing the file using Microsoft’s in-built utility - Open and Repair. To use this utility, follow these steps:

- In Excel, go to **File > Open > Browse**.
- In the Open dialog box, click on the corrupted Excel file.
- Click the arrow next to the Open button and select **Open and Repair** from the dropdown.
- Select **Repair** to recover as much data from the file as possible.

If the Open and Repair utility fails or stops working, then you can try a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can repair severely corrupted Excel files **(.xls, .xlsx, .xltm, .xltx, and .xlsm)**. It helps recover all the file components, including images, charts, tables, pivot tables, cell comments, chart sheets, formulas, etc., without impacting the original structure.

## **Conclusion**

The Runtime error 424 usually occurs when there is an issue with the objects in your VBA code. In this article, we have covered some effective methods to resolve the “object required” error in Excel. If the error occurs due to corruption in Excel file, then you can repair the corrupt file using Stellar Repair for Excel. It is a reliable tool that can repair severely corrupted Excel file without changing its actual formatting. You can download the free trial version of the software to evaluate its functionality.


## Solved - The File is Corrupted and Cannot be Opened - Excel

**Summary:** Unable to open Excel file due to the error ‘The file is corrupted and cannot be opened’? Read this blog to find more details about the error, possible reasons behind it, and solutions to fix the error. In addition, the blog mentions about Stellar Repair for Excel software that can help fix the Excel error in a few clicks. Download the software now and see free preview of the file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

## **About the Error**

**Microsoft Excel** is a widely used spreadsheet application that comes bundled with MS Office. Users tend to update the application with new security patches and features. Sometimes these updates can cause problems, and result in “**The file is corrupted and cannot be opened**” error.

![The File is Corrupt and Cannot be Opened Error Message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/the-file-is-corrupt-and-cannot-be-opened.jpg)

Figure 1 – Excel File Corrupted Error Message

## **Other Possible Reasons behind ‘The File is Corrupt and Cannot Be Opened’ Excel Error**

- Opening an older Excel version file in a newer version of Excel. For instance, opening Excel 2013, 2010, or earlier versions in Excel 2016.
- When attempting to open a Microsoft Office (Excel) email attachment in Microsoft Outlook 2010, MS Office 2010 reports a problem with the file preventing it from opening.

## ****How to Fix the ‘Excel File is Corrupt and Cannot Be Opened’ Error?****

Here are a few possible solutions that you can try to fix the ‘Excel file is corrupt and cannot be opened’ issue and open your Excel file.

**Solution 1**: Changing Component Services Settings

**Solution 2**: Changing the Protected View Settings

**Solution 3**: Repair Excel Files using Excel Repair Software

### **Solution 1: Changing Component Services Settings**

**\[Caution\]** Changing Component Services settings requires making changes to the registry, and any mistake can harm your computer.

**Follow these steps to change ‘Component Services’ settings:**

- Click ‘**Start**’ or ‘**Win+R**’ and type ‘**dcomcnfg**’ and press ‘**Enter’.**

- In the navigation pane, expand the ‘**Component Services**’, and then expand ‘**Computers**’.

___

![Changing Component Services Settings](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/Changing-Component-Services-Settings.jpg)

Figure 2 – Component Services Settings

- Next, right-click on ‘**My Computer’**, and then click ‘**Properties**’.

**When the ‘My Computer Properties’ dialog box appears, click on the ‘Default Properties’ tab and then set the following values:**

- **Default Authentication Level**: Connect
- **Default Impersonation Level**: Identify

![My Computer Properties](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/My-Computer-Properties.jpg)

Figure 3 – Illustrates My Computer Properties

- Click ‘**OK**’ to change ‘**Default Properties**’

### **Solution 2: Changing the Protected View Settings**

**\[Caution\]** Disabling the ‘Protected View’ can put your system at high risk. Viruses attached to the Excel files can attack and infect your system. Be careful before using this option.

Excel 2010 file cannot open due to the ‘**Protected View**’ setting in Microsoft Outlook 2010. And so, changing the setting may help fix the error. For this, perform these steps:

- Open MS Excel 2010, go to the ‘**File’** menu and click **‘Options’.**

![Select Options in Excel 2010](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-options.jpg)

Figure 4 – Options

- When the ‘Excel Options’ window opens, click on ‘**Trust Center**’ and then on ‘**Trust Center Settings**’.

![Trust center settings in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Open-trust-center-settings.jpg)

Figure 5 – Open Trust Center Settings

- Next, choose **‘Protected View**’ and uncheck all the options including ‘**Enable Protected View for Outlook attachments’** if you use Outlook for email.

![change protected view settings](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Uncheck-protected-view-settings.jpg)

Figure 6 – Uncheck Protected View Settings

- Click ‘**OK’.** Restart the application and try opening the Excel file again.

If none of the above solutions works for you, your Excel file is likely severely corrupt. To repair corrupt Excel files, you need to use advanced options like Stellar Repair for Excel tool. It repairs corrupt and damaged Excel files and helps in retrieving lost data.

### **Solution 3: **Use Excel File Repair Tool****

Considering the risks associated with the above solutions, it’s better to use an **Excel repair tool** to repair **single** or **multiple** corrupt Excel files at once. The process is simple, and even a novice can use the Excel file repair tool to repair Excel files with the help of the following steps:

- Download Stellar Repair for Excel and install it.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

- Launch the tool. In the tool’s main interface, click ‘**Browse**’ to select the file. If you don’t know the file location use the ‘**Search’** option.

![Browse and Search](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/Browse-and-Search.jpg)

Figure 7 – Illustrates Selecting Corrupt Excel File in Stellar Repair for Excel

- Select the file, and then click on **Repair**.

![select corrupt file and repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/2-select-file.jpg)

Figure 8 – Illustrates Initiating Excel File Repair in Stellar Repair for Excel

- The software scans and lists the Excel file in the left pane. Click on the file to preview its recoverable objects in the right pane.

![preview recoverable excel objects](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/4-preview.png)

Figure 9 – Illustrates Preview of Recoverable Excel File Objects

- Save the repaired file at either the default location or a user-specified location.

![select repaired file location](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/6-save-file.jpg)

Figure 10 – Illustrates Saving Repaired Excel File in Stellar Repair for Excel

- Click ‘**OK’** to save the repaired Excel file. After the repair process is completed, browse to the location and open it with MS Excel 2010 or any other version.

![repaired file saved Dialog Box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/7-saving-complete.jpg)

Figure 11 – Illustrates Saving Complete Message in Stellar Repair for Excel

You will be able to access your Excel file from the selected location.

## **Conclusion**

You can use the first two possible solutions to fix the “The file is corrupted and cannot be opened” error. If you can access the file, save its data and restore the default settings. However, if the file is corrupt and the data retrieved using the first two solutions is inconsistent or incomplete, use Stellar Repair for Excel. This tool can help you recover Tables, Charts, Chart Sheets, cell comments, Images, and Formulas while preserving the worksheet properties and cell formatting. You can also preview the file and verify the data inside the file before saving it.



## How to Fix Microsoft Excel Error Code 0x800A03EC?

**Summary:** You can encounter the error code 0x800A03EC in Excel due to different reasons. This post discusses the causes of the error and the workarounds and methods to fix it. If the “Microsoft Excel error 0x800A03EC” prevents you from accessing the Excel file data, use the Excel repair file tool mentioned in this post to restore the data with complete integrity.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

The error code 0x800A03EC in MS Excel can occur while exporting or importing Excel sheets to or from another application. It can occur when trying to execute a macro. It usually appears when the Excel application fails to read the queries in the VBA code, especially when using the PasteSpecial method (Range.PasteSpecial method) to paste cell data from the clipboard into an Excel sheet. It can occur if the cell data format is not compatible with the target data.

## **Causes of MS Excel Error Code 0x800A03EC**

The error code 0x800A03EC in Excel can occur in different scenarios. There could be several reasons associated with this error. Some of them are:

- Incompatible data formats.
- Trying to paste an extensive range of cells into a smaller range of cells.
- Add-ins are interrupting the code operation.
- Issue with the VBA code (incorrect or incomplete queries).
- Excel file is corrupted/damaged.
- Corrupted macros.
- Trying to export large-sized Excel file which is more than the Excel’s prescribed limit.
- Outdated Excel version.
- Incompatible cell formats on source and destination (when copying/pasting cell data).
- Missing cell range (forget to specify the range while using PasteSpecial method in the VBA code).
- Trying to save Excel file to incorrect directory.
- File path contains invalid characters.

## **Solutions to Fix MS Excel Error Code 0x800A03EC**

The error 0x800A03EC can appear if your Excel file is incompatible with your Excel application version. You can run the [compatibility checker](https://support.microsoft.com/en-us/office/save-an-excel-workbook-for-compatibility-with-earlier-versions-of-excel-169a0336-965b-4430-8554-4e7b5db79947) to review the compatibility issues in your Excel file. If this is not the issue, then follow the below methods.

### Method 1: Review VBA Code

The exception from hresult 0x800a03ec excel can occur if you are trying to execute incorrect or incomplete queries in the Excel VBA code. It can also appear if the formulas do not contain the equal (=) symbol at the beginning. Verify the VBA code for any logical flaws, typo errors, syntax errors, or missing references.

### Method 2: Check the Field Size

Excel has certain [limits and specifications](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us). If the Excel file’s data exceeds these designated size limits, you can get the MS Excel error code 0x800A03EC error. For example, this error occurs if you try to export a file with more rows than the limit of 65536 and columns than the limit of 256. Check and optimize the file size by minimizing complex formulas and other objects.

### Method 3: Check Add-ins for Disabled Items

Sometimes, disabled items in Add-ins settings can prevent macros from functioning correctly. You can check and enable the disabled items in Add-ins using these steps:

- Navigate to **File > Options**.

![Go To Options Tab](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/08/go-to-options-tab.jpg)

- In **Excel Options**, click on the **Add-ins** option.
- Click the arrow corresponding to the **Manage** section.
- Select **Disabled Items** and click on the **Go** option.

![Go To Addins And Select Disabled Items](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/08/go-to-addins-and-select-disabled-items.jpg)

- You will see a list of disabled items.
- Click on the disabled items and then click **Enable**.
- Restart Excel for the applied changes to take effect.

### Method 4: Change Macro Settings

The error code 0x800A03EC can also occur if macros are disabled in the Macro Security settings. Follow these steps to change the macro settings in Excel:

- In MS Excel, go to **File > Options > Trust Center**.

![Click Trust Center Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/08/click-trust-center-option.jpg)

- Under Trust Center, click on **Macro Settings**.

![Click Macro Settings and Selecting Enable All Macros](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/08/click-macro-settings-select-enable-all-macros.jpg)

- Select **Enable all macros** and click **OK**.

### Method 5: Check the OLE Objects

The Microsoft Excel error code 0x800A03EC can also appear if there is a connection disruption in the VBA. Such an issue can occur if there is an issue with Object Linking and Embedding (OLE) in an Excel workbook. The OLE objects are linked to external files. You can check and remove the unnecessary OLE objects from your Excel file to fix the issue.

### Method 6: Check Methods in VBA

Excel can throw the “Exception from HRESULT: 0x800A03EC” error if you are trying to call an invalid method in a VBA code. Many users have reported this issue when trying to use ‘copy and paste’ feature using copy paste special [method](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial) (range.pastespecial) in Excel. Paste special is an advanced option in Excel to smooth the copy-and-paste task. While using this method, the exception can usually occur when Excel application detects an invalid or misaligned range or mismatch data type in the syntax. To fix this, check the syntax of the paste special method.

### Method 7: Repair your Excel File

Corruption in Excel file can create inconsistencies in the macro and lead to the “Exception from HRESULT 0x800a03ec Excel” error. In such a case, you can try repairing the Excel file using Microsoft’s inbuilt utility – Open and Repair. To use this utility, follow these steps:

- In the Excel application, go to the **File** tab and then click **Open**.
- Click **Browse** to select the Excel file in which you are getting this exception error.
- The **Open** dialog box will appear. Click on the corrupted file.
- Click the **arrow** next to the Open button and then select **Open and Repair**.
- You will see a dialog box with three buttons – Repair, Extract Data, and Cancel.

![Click On Repair Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/08/click-on-repair-option.jpg)

- Click on the **Repair** button to recover as much of the data as possible.
- After repair, a message is displayed. Click Close.

The Open and Repair tool works in significant scenarios. If [Open and Repair tool fails](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to resolve the issue, try an Excel repair tool recommended by experts and MVPs to repair the corrupted Excel file. Stellar Repair for Excel is one of the recommended Excel repair tools for fixing issues caused by corrupt Excel files. It provides more features than the Open and Repair utility. It can even repair severely corrupted Excel files. The tool supports all the versions of Microsoft Excel, including 2019. Download the software’s demo version to scan the corrupted file and see the preview of all the recoverable components of the file.

### Conclusion

There are numerous reasons, like invalid method, incorrect range, data type mismatch, etc., that could lead to the common exception error 0X800A03EC in Excel. Try the troubleshooting methods mentioned above to fix the issue. You can also encounter the error due to corruption in the Excel file. In such a case, you can try the professional software – [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) – to repair the severely corrupted Excel file. The Excel repair software can fix all the corruption-related issues in Excel files (XLS/XLSX).



## How to Repair Corrupt Excel Workbook?

**Summary:** Corruption in an Excel workbook could lead to data loss. This blog outlines the possible reasons behind corrupted workbook and describes how to repair the workbook manually and using an Excel repair tool.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Is your Excel behaving differently, for instance, [stops responding](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) or freezes? Or have you encountered a warning or an error message indicating corruption in the Excel workbook?

![Corruption in Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Corrupt-workbook-cannot-be-opened-or-repaired-1.png)

Figure 1 - Corruption Warning in Excel: The corrupt workbook cannot be opened or repaired

![Excel file Corruption Message](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Excel-file-corruption-message-2.png)

Figure 2 - Excel File Corruption Message

In any of these situations, you won?t be able to open your Excel workbook.

**_TIP!_** _If you?re unable to open an Excel workbook, it does not necessarily mean that the workbook is corrupt. Try a few workarounds to resolve the issue, like disabling the ?Protected View Settings?, disabling add-ins, or moving the workbook to another drive._

You will need to repair the corrupted Excel workbook to make it usable. Before discussing the repair methods, let?s first look at the possible reasons behind file corruption.

**To jump to methods to repair the workbook, [click here](https://www.stellarinfo.com/blog/repair-a-corrupted-workbook/#_Methods_to_Repair).**

## **Possible Reasons behind Corrupted Excel Workbook**

- **<u>Large size of a workbook</u>**: A workbook containing lots of information or objects like images, shaded cells, or other visual representations is prone to corruption.
- **<u>Unexpected system shutdown or power loss</u>**: If you haven?t saved your Excel file, abrupt system shutdown or power failure can corrupt the data and render the file inaccessible.
- **<u>Bad sectors on storage media</u>**: Accumulation of bad sectors on storage media on which Excel files are saved is another reason that may lead to Excel file corruption.

Other common reasons causing Excel file corruption are virus attack, network errors, etc.

## **Methods to Repair Corrupted Excel Workbook**

**_Note:_** _Repairing the corrupt workbook may result in data loss. And so, it?s recommended to back up the workbook before attempting to repair the file._

### **Method 1 ? Repair the Excel Workbook Manually**  

Use the Excel ?Open and Repair? feature to repair your corrupt workbook manually. Here?s how:

Step 1: In Excel, click the **File** tab and then click **Open**.

Step 2: Select the corrupt Excel file, click the arrow button next to the Open button and then select **Open and Repair**.

![Excel Open and Repair Feature](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Excel-open-and-repair-feature-3.png)

Figure 3 - Excel Open and Repair Feature

Step 3: In Microsoft Excel dialog box that opens, click **Repair** to begin repairing your corrupt workbook.

![Repair the Microsoft Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Repair-Excel-workbook-4.png)

Figure 4 - Repair the Microsoft Excel File

**_Note_**_: If the Excel repair process fails, repeat Steps 1 till 3 and click the ?Extract Data? button. This will help you extract values and formulas from the corrupted workbook._

Performing these steps will repair your Excel file. If the issue persists, skip to the next method.

### **Method 2 ? Use Excel File Repair Software**

If the above method failed to fix corruption in the workbook, use [Excel file repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to repair your workbook quickly. This Excel repair tool is purpose-built to help users repair large-sized Excel files (.xlsx and .xls) in just a few clicks. Also, it helps recover all the data from the corrupted workbook while preserving the cell formatting and worksheet properties.

## **Steps to Repair Corrupt Excel File Using the Software**

Step 1: Download, install, and open **Stellar Repair for Excel** software.

[Free Download for windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Step 2: Click **Browse** to select your corrupted .xlsx or .xls file. If you are not aware of the file location, click **Search** to locate and select the file.

![Select Corrupt Excel Workbook](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Select-corrupt-excel-workbook-5-1.png)

Figure 5 - Select Corrupt Excel Workbook

Step 3: Once the file is selected, click **Repair**.

Step 4: A preview of recoverable workbook data gets displayed.

![Preview of Recoverable Excel Workbook Data](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Preview-recoverable-excel-workbook-data-6-1024x671.png)

Figure 6 - Preview of Recoverable Excel Workbook Data

**_Note:_** _You will need to activate the licensed version of Stellar Repair for Excel software for saving the repaired file. Once activated, proceed with the steps below to save the repaired file._

Step 5: Click **Save File** on the **File** menu. This will open a **Save File** dialog box with options to save the repaired .xlsx/.xls file at default or new location. After choosing the appropriate option, click **OK**.

![Save Repaired Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/03/Save-repaired-file-7.png)

Figure 7 - Save Repaired Excel File

The repaired file will get saved at the selected location.

##### **Key Features of Stellar Repair for Excel Software**

- Repairs severely corrupt Excel file and recovers pivots, pivot tables, conditional formatting rules, chart, chart sheets, etc.
- Can repair multiple Excel files in a batch simultaneously
- Can handle all types of Excel file corruption errors, like ?Excel found unreadable content in filename.xls?, ?Excel cannot open the file filename.xlsx?, etc.
- Supports repairing corrupt workbooks of Excel 2019, 2016, 2013, 2007, 2003, and lower versions.

## **End Note**

If your Excel workbook has turned corrupt, try using the methods discussed in this blog to repair the workbook. Once it is repaired, you must follow preventive measures to avoid your Excel file from getting corrupt. Some of these measures are as follows:

- [Reduce the size of your workbook](https://support.microsoft.com/en-us/topic/reduce-the-file-size-of-your-excel-spreadsheets-c4f69e3a-8eea-4e9d-8ded-0ac301192bf9#ID0EBDD=Office_2013_-_2016) by:  
  - Removing unused data
  - Saving images at lower resolutions
  - Avoid saving pivot table cache

- Turn on the ?AutoRecover? and ?AutoSave? features in Excel





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


## How to Fix a Corrupted .xls File? The Everything Guide

Undoubtedly, Excel is so powerful that it can help you to process, analysis, and store data, in masses.

That’s the reason it has been there for years and helping this world in data.

But…

With all those powers comes some nasty problems which no Excel users like to face. Can you guess what I’m talking about?

Think about a Corrupted Excel File. Nightmare? Isn’t it?

And do you remember that last time when you have opened a workbook and you got a message that this workbook is might corrupt?

The TRUTH is, this is something which you cannot avoid, but, you can [prepare yourself in the best way](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) and deal with it like a PRO.

So today, in this post, I’d like to share with you to everything you need to know about a corrupt Excel file (.xls), why it happens, how to fix it like a PRO, and much more.

...let’s get started.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/introduction.jpg)

**Note**: In this post, we’ll be covering the .xls version (which is the extension for the file which is created in Excel 2007 or the earlier versions) and if you want to know about the new version, [here’s the quick fix](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) for that.

## Why My Excel File Got Corrupted?

There can be one or multiple reasons for an Excel file to get corrupted. Below I have detailed about some of the major of them.

### 1\. Large Excel File

You can store data in a workbook the way you want but sometimes using excessive thing can make an Excel file bigger in size.

And that kind of data files can crash at any point in time. Here are a few things which make the Excel files heavy, like

- Conditional Formatting.
- Colors formatting.
- Using merged cells in place of text alignment.
- Volatile functions: Formulae that iterate every time you open or change a cell value; OFFSET, NOW.
- Using a complete column or row as a reference than the data set range.
- Using complex formulas; VLOOKUP in place of Index/Match, Nested If in place of MAXIFS, MINIFS.
- Calculations or reference across workbooks.

**Related:** [How to Fix Formatting Issues in Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

### 2\. Abrupt System Shutdown

Shutting down the system without following the procedure can corrupt your data file.

This shut down can be due to a power failure or any other unexpected technical challenges.

So it is always important to follow the procedures and shut down your system properly to avoid data losses.

### 3\. Infected Excel File (Virus Attack)

This is the most common and obvious reason for Excel file corruption.

Although we always keep our system safe using various Antiviruses, still there is always a probability of virus attacks and loss of important files.

It is always advised to use a safe and strong antivirus compatible with your system requirements.

## What are the Signs to Know When an Excel File is Corrupted?

In this section, we will discuss what are the signs which you can get when an Excel file is corrupted, let’s dig into it.

### 1\. The File is Corrupt and Cannot Be Opened

This is one of the most common messages you can see when your workbook is corrupted.

But there is also a chance that it is just because of the version compatibility where you have a .xls file but you are using the latest version of Excel check out this detailed post by Priyanka

### 2\. We Found a Problem with some Content in this File…

There’s another error message which you can get while opening a file:

We Found a Problem with some content in Do you want us to recover as much as we can? If you trust the source of this workbook, click yes.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/we-found-a-problem-with-some-content-in-this-file.jpg)

There are a lot of applications out there (I think almost every) which exports the data as a .xls format. Those files have a greater chance of having this kind of error.

### 3\. “Filename.xls” cannot be accessed

There can also be a situation where you get the error:

_“Filename.xls” cannot be accessed. The file may be corrupted, located on a server that is not responding._

Well, this message is a bit misleading.

You won't be able to decide that your file is actually corrupted or just not on the location.

## My Excel File Got Corrupted, now What Should I Do?

There are many ways to recover the data from the corrupt excel files. But before you start, it is always advised to create a copy of the corrupted file.

You can save a lot of time with [**Stellar Repair for Excel,**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) which make data recovery just with few clicks.

But before you go for a data recovery software, let's try out some manual steps which can help.

When a workbook get corrupted the first thing comes to the mind is to recover data from it…

...and you what there’s a simple option there in the Excel which you can use to do this. Below are the steps you need to follow:

- First of all, open the Excel and click on the office icon.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/office-icon.jpg)

- After that, go to the “Open” and select the file which is corrupted.

- Now, click on the open drop-down and select “Open and Repair”.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repair.jpg)

- At this point, you have two options:

1. **Repair File**
2. **Extract Data**

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/extract-data.jpg)

Let’s get into both of these options one by one...

### 1\. Repair File

This option helps you to repair the file and the moment you click on it it takes a few seconds afterward and shows you the result with a message box and also provide you a log file.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/repair-file.jpg)

And once it is done with repairing, you'll get your file opened and you can save that file as a new copy.

Yes, that’s it.

### 2\. Extract Data

If somehow you aren’t able to get your file repaired, you can also extract data from that file using “Extract Data” option.

Even in this option, you can get data in two ways.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repair-options.png)

1. **As Values**
2. **With Formulas**

In the first option, Excel simply extracts data as value ignoring all the formulas driving those value (which is **the best way if you just need to have that data back**).

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repai-values.jpg)

But in the second option, Excel tries to recover the formulas as much as possible.

Check out this [**smart technique by Jyoti**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) which you can use it you aren’t able to recover data from the file.

## Preventions to Not to have any Excel File Go Corrupt in Future

Future is fragile, what I’m trying to say is the more you work in Excel and process data there could be a chance that your workbook goes corrupt.

If there’s no security then what an EXCEL POWER user should do?

Well, there are few things which you can do or take care of while working with Excel so that you won’t have to worry about corruption of Excel workbooks.

Let’s see what you can do…

### 1\. Change Recalculation Option

Now here’s the thing when you work with a hell lot of data, there a common thing that you gotta using formulas. Right?

But, the thing these formulas are something which makes your Excel file slows down sometimes make them go corrupt.

There’s one small tweak you can do in your workbook is change the calculation method.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/change-recalculation-option.jpg)

Now with the manual calculation, you just need to whenever you open your file it won’t recalculate all the formulas.

And when you update your data you can simply click on the “Calculate Now” and it will calculate all the formulas again.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/calculate-now.jpg)

**Quick Tip:** Beware of Volatile Functions and use them with caution as recalculates them every time you change something in the worksheet.

### 2\. Use VBA Codes Instead of Formulas

Now, this is what I do when I need to use complex formulas in a workbook.

Here’s how you can do this: Let’s say you have a formula in the cell A1, like below, which calculates the age.

\="You age is "& DATEDIF(Date-of-Birth,TODAY(),"y") &" Year(s), "& DATEDIF(Date-of-Birth,TODAY(),"ym")& " Month(s) & "& DATEDIF(Date-of-Birth,TODAY(),"md")& " Day(s)."

Now, instead of simply entering it into the cell A1 which I would write a macro code which inserts this formula into the cell A1 and then convert it into the a value.

**Here’s the code:**

Sub CalculateAge()  
Range("B1").Value = \_  
"=""Your age is """ & \_  
"&DATEDIF(A1,TODAY(),""y"")" & \_  
"&"" Year(s), """ & \_  
"&DATEDIF(A1,TODAY(),""ym"")" & \_  
"&"" Month(s), and """ & \_  
"&DATEDIF(A1,TODAY(),""md"")" & \_  
"&"" Days(s)."""  
Range("B1") = Range("B1").Value  
End Sub

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/VBA-codes.jpg)

**Note:** To write these code you need to have basic understading of VBA (make sure [check out this guide](https://excelchamps.com/learn-vba/) for this).

### 3\. Use a File Recovery Application

Recently we asked a quick question to our readers on ExcelChamps that if they have ever faced a situation where they got a corruption message in Excel.

You’ll be astonied to hear that 50% percent of the people said “YES” they faced this thing in the past.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/yesr.jpg)

Now, this is alarming, if you are heading a team or you have a bunch of people in your company who use Excel…

…there’s a high probability that half of them gonna face this issue. So the best way to deal with this to have an App FIX your Excel file for you.

With **STELLAR REPAIR FOR EXCEL,** you just need a few clicks, yes that’s right. Let me show you with the below steps:

- First of all, download the app and install it (it’s simple).

[![download](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

- After that, open the app and click on the “Browse” and simply select the file which is corrupted.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-home.jpg)

- In the end, click on the REPAIR to let the Excel repair software fix your file (it takes a few seconds).

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-log-report.jpg)

Once you complete repairing your file, you’ll get a message in your on the status bar and after that, you can open your file.

## Final Thoughts

If you are a POWER Excel user then there’s a must for you to have known how to deal with a situation where you got a corrupt Excel file.

But I must recommend you to TRY OUT Stellar Repair for Excel so that’s you don’t have to worry about your Excel files anymore.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-review.jpg)

I’m sure you found this post helpful, and please don’t forget to share this tip with your colleagues, I’m sure they’ll appreciate it.


## How to Fix the Unable to Record Macro Error in Excel?

**Summary:** You may encounter the “Unable to record macro” error in MS Excel when using Personal Macro Workbooks. In this post, we’ll discuss the possible causes behind this error and the ways to fix it. We’ll also mention a professional Excel repair tool that can help fix the error if it occurs due to corrupted workbook.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

A personal macro workbook (Personal.xlsb file) is a hidden workbook in Excel that stores all macros in a single workbook. This allows you to automate processes while working in Excel. Sometimes, when recording macro codes in the personal macro workbook, you may get the message “**Personal Macro Workbook in a startup folder must stay open for recording**”. When you click on the **OK** button, it will show the “unable to record” error. This prevents you from recording the macros. Below, we’ll see the causes behind this error and discuss how to resolve this error.

## **Causes of Unable to Record Macro Error**

You may be unable to record macros in Excel due to several reasons. Let’s take a look at the possible causes that can lead to this issue.

- The location of personal.xlsb file is changed.
- Personal.xlsb file is corrupted.
- Macros are disabled.

## **Methods to Fix the “Unable to Record Macro” Error in Excel**

Here are some possible solutions that can help you resolve the unable to record macro error in Excel.

### Method 1: Check the Path of XLStart Folder

You may be unable to record macros if the path of XLStart folder is incorrect. It is a folder where the Personal.xlsb file is stored by default. Follow these steps to find out the path of this folder:

- Open MS Excel. Go to **File > Options**.
- Click **Trust Center > Trust Center Settings**.

![Excel Options Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-window.png)

- In the **Trust Center Settings** window, click on **Trusted Locations**.

![Path Of XLStart Folder In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/path-of-xlstart-folder-in-trust-center.png)

- Verify the path of the **XLSTART** folder and modify it if there is an issue.
- Once you are done, click on **OK**.

### Method 2: Change Macro Security

The “Unable to record macro” error can occur if macros are disabled in the Macro Security settings. You can try changing the macro settings using the below steps:

- In MS Excel, go to **File > Options > Trust Center**.

![Excel Options To Locate Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-to-locate-trust-center.png)

- Under **Trust Center,** click on **Trust Center Settings**.

![Change Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/change-macro-settings-in-trust-center.png)

- Select **“Enable all macros”** and then click **OK.**

### Method 3: Check Add-ins for Disabled Items

If there are any items in add-ins that are disabled, they may prevent Excel from functioning properly. You can check and enable the items in MS Excel using the below steps:

- Click **File > Options.**

![Go To Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/go-to-options-1.png)

- In **Excel Options**, click on the **Add-ins** option.
- Select **Disabled Items** from the **Manage** section and click on **Go**.

![Add-ins In Excel Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/add-ins-in-excel-options.png)

- The **Disabled Items** window is displayed.
- Click on the disabled item and then click **Enable**.
- Restart Excel for the changes to take place.

### Method 4: Repair your Excel File

You may fail to record macros if there is corruption in the workbook. In such a case, you can use the “Open and Repair” utility in MS Excel to repair the corrupt workbook. To use this tool, follow these steps:

- Open your Excel application.
- Click **File > Open**.
- Browse to the location where the corrupted file is stored.
- In the **Open** dialog box, choose the corrupted workbook.

![Open Dialog Box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-dialog-box.png)

- From the **Open** dropdown list, click **Open and Repair**.

![Open And Repair Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-and-repair-window.png)

Excel will prompt you to repair the file or extract data from it. Click **Repair** to retrieve maximum data. If the Repair option fails, then click on the Extract Data option to recover the data without formulas and values.

If the Microsoft utility “Open and Repair” fails to repair the corrupted Excel workbook, then try a professional Excel repair tool such as Stellar Repair for Excel. It is an advanced tool that can easily repair severely corrupted Excel (XLSX and XLS) files. It can recover all the file items, including chart sheets, cell comments, tables, macros, formulas, etc. without impacting the properties and cell format of the Excel file.

## **Closure**

You may receive the “unable to record” error in Excel while creating or storing macros in Personal Macro Workbooks. There are several reasons that can lead to this error. You can try the methods covered in this post to resolve the error. If the error appears due to corruption in workbook, then try to repair it using the Open and Repair utility. Alternatively, you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) – a professional tool that can help you repair an Excel file with problematic macros. Also, it allows recovery of all the file components with complete integrity. The tool is compatible with Excel 2021, 2019, 2016, and older versions.


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
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-infinix-smart-8-hd-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Infinix Smart 8 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-8-without-losing-data-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone 8 without Losing Data? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-13-pro-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 13 Pro Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-lost-or-stolen-iphone-7-in-easy-steps-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Data From Lost or Stolen iPhone 7 In Easy Steps | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-ios-system-issues-of-iphone-13-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iOS System Issues of iPhone 13 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-htc-u23-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of HTC U23 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-honor-play-8t-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Honor Play 8T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-oneplus-open-by-fonelab-android-recover-data/"><u>How to recover lost data from OnePlus Open?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-12-pro-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 12 Pro Data From iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-microsoft-excel-2003-has-stopped-working-error-stellar-by-stellar-guide/"><u>How to fix Microsoft Excel 2003 has stopped working error? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-vivo-y28-5g-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Vivo Y28 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-oneplus-nord-ce-3-5g-by-fonelab-android-recover-data/"><u>How to retrieve lost files from OnePlus Nord CE 3 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-vivo-y02t-by-fonelab-android-recover-music/"><u>How to recover old music from your Vivo Y02T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-motorola-razr-40-ultra-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Motorola Razr 40 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-oppo-f23-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Oppo F23 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-infinix-note-30-pro-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Infinix Note 30 Pro to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-v29-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from V29</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-infinix-zero-30-5g-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Infinix Zero 30 5G Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-vivo-y36i-by-fonelab-android-recover-video/"><u>How to recover old videos from your Vivo Y36i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-htc-u23-pro-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove HTC U23 Pro PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-8-without-losing-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 8 without Losing Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-value-error-in-excel-2003-stellar-by-stellar-guide/"><u>How to Fix the #Value! Error in Excel 2003? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-drivers-with-windows-device-manager-in-windows-1110-by-drivereasy-guide/"><u>How to identify malfunctioning drivers with Windows Device Manager in Windows 11/10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-an-mp4-on-motorola-edge-40-neo-by-aiseesoft-video-converter-play-mp4-on-android/"><u>How to play an MP4 on Motorola Edge 40 Neo?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-nokia-150-2023-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Nokia 150 (2023) using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-k11-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from K11 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-nubia-red-magic-9-proplus-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Nubia Red Magic 9 Pro+.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-oppo-find-n3-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Oppo Find N3.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-12-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi 12</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-erase-private-data-from-iphone-15-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Erase Private Data From iPhone 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-find-x6-by-fonelab-android-recover-data/"><u>How to recover lost data from Find X6?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-honor-90-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Honor 90 using Video Repair Utility?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-pdf-v14-document-with-digital-signature-software-by-ldigisigner-sign-a-pdf-sign-a-pdf/"><u>How to sign a PDF v1.4 document with digital signature software</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-itel-a60s-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Itel A60s Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-your-hardware-drivers-with-windows-device-manager-in-windows-10-and-7-by-drivereasy-guide/"><u>How to identify missing your hardware drivers with Windows Device Manager in Windows 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-x-to-an-older-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone X to an Older iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6s-plus-without-losing-any-content-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6s Plus without Losing Any Content? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-15-pro-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 15 Pro without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-excel-2000-formulas-not-working-properly-step-by-step-guide-stellar-by-stellar-guide/"><u>How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-itel-s23plus-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Itel S23+ to iPhone | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/insert-signature-in-wbk-file-by-ldigisigner-sign-a-word-sign-a-word/"><u>Insert signature in .wbk file</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-resolve-your-apple-iphone-11-pro-max-keeps-asking-for-outlook-password-drfone-by-drfone-ios/"><u>In 2024, Resolve Your Apple iPhone 11 Pro Max Keeps Asking for Outlook Password | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/how-to-unlock-apple-id-activation-lock-on-iphone-12-by-drfone-ios/"><u>How to Unlock Apple ID Activation Lock On iPhone 12?</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/about-tecno-frp-bypass-by-drfone-android/"><u>About Tecno FRP Bypass</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-locked-out-of-apple-iphone-13-pro-5-ways-to-get-into-a-locked-apple-iphone-13-pro-by-drfone-ios/"><u>In 2024, Locked Out of Apple iPhone 13 Pro? 5 Ways to get into a Locked Apple iPhone 13 Pro</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-3-easy-ways-to-factory-reset-a-locked-iphone-11-pro-without-itunes-by-drfone-ios/"><u>In 2024, 3 Easy Ways to Factory Reset a Locked iPhone 11 Pro Without iTunes</u></a></li>
<li><a href="https://android-frp.techidaily.com/hassle-free-ways-to-remove-frp-lock-on-samsung-galaxy-a23-5gwithwithout-a-pc-by-drfone-android/"><u>Hassle-Free Ways to Remove FRP Lock on Samsung Galaxy A23 5Gwith/without a PC</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-exploring-the-best-slow-motion-app/"><u>New Exploring the Best Slow-Motion App</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-apple-id-locked-for-security-reasons-from-apple-iphone-13-mini-find-the-best-solution-here-by-drfone-ios/"><u>In 2024, Apple ID Locked for Security Reasons From Apple iPhone 13 mini? Find the Best Solution Here</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/text-to-speech-online/"><u>Text To Speech | Online</u></a></li>
<li><a href="https://android-frp.techidaily.com/in-2024-how-to-bypass-google-frp-lock-on-oppo-a78-devices-by-drfone-android/"><u>In 2024, How to Bypass Google FRP Lock on Oppo A78 Devices</u></a></li>
<li><a href="https://fake-location.techidaily.com/spoofing-life360-how-to-do-it-on-motorola-edgeplus-2023-drfone-by-drfone-virtual-android/"><u>Spoofing Life360 How to Do it on Motorola Edge+ (2023)? | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/updated-in-2024-windows-movie-maker-download-tutorial-from-installation-to-video-editing-mastery/"><u>Updated In 2024, Windows Movie Maker Download Tutorial From Installation to Video Editing Mastery</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-videos-from-lava-yuva-3-pro-by-fonelab-android-recover-video/"><u>Easy steps to recover deleted videos from Lava Yuva 3 Pro</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/the-ultimate-guide-how-to-bypass-swipe-screen-to-unlock-on-realme-narzo-60-5g-device-by-drfone-android/"><u>The Ultimate Guide How to Bypass Swipe Screen to Unlock on Realme Narzo 60 5G Device</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/how-can-i-get-more-stardust-in-pokemon-go-on-realme-v30t-drfone-by-drfone-virtual-android/"><u>How can I get more stardust in pokemon go On Realme V30T? | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-how-can-i-catch-the-regional-pokemon-without-traveling-on-honor-magic-vs-2-drfone-by-drfone-virtual-android/"><u>In 2024, How Can I Catch the Regional Pokémon without Traveling On Honor Magic Vs 2 | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/open-source-video-editing-on-a-budget-top-free-options-for-2024/"><u>Open-Source Video Editing on a Budget Top Free Options for 2024</u></a></li>
<li><a href="https://fake-location.techidaily.com/full-guide-to-fix-itoolab-anygo-not-working-on-tecno-pop-8-drfone-by-drfone-virtual-android/"><u>Full Guide to Fix iToolab AnyGO Not Working On Tecno Pop 8 | Dr.fone</u></a></li>
</ul></div>



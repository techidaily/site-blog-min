---
title: How to Fix Excel 2007 Run Time Error 1004 | Stellar
date: 2024-06-18T01:01:00.957Z
updated: 2024-06-19T01:01:00.957Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2007 Run Time Error 1004
excerpt: This article describes How to Fix Excel 2007 Run Time Error 1004
keywords: repair .xlsm files,repair damaged .xlb,repair corrupt .csv files,repair damaged .xlsm files,repair damaged .xlb files,repair damaged .csv,repair excel 2013,repair excel 2000,repair .xltx
thumbnail: https://thmb.techidaily.com/64bcba811dca59ee452fde50283dc6af9516c46b5a87dc01f6fa89f4e4093f9a.jpg
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

### **Create a Macro to Reduce Chart Legend Font Size**

Sometimes, Excel throws the run-time error when you try to run VBA macro to change the legend entries in a Microsoft Excel chart. This error usually occurs when Microsoft Excel truncates the legend entries because of the more legend entries and less space availability. To fix this, try to create a macro that shrinks/minimize the font size of the Excel chart legend text before the VBA macro, and then restore the font size of the chart legend. Here is the macro code:

```
VBCopy
Sub ResizeLegendEntries()

With Worksheets("Sheet1").ChartObjects(1).Activate
      ' Store the current font size
      fntSZ = ActiveChart.Legend.Font.Size

'Temporarily change the font size.
      ActiveChart.Legend.Font.Size = 2

'Place your LegendEntries macro code here to make
         'the changes that you want to the chart legend.

' Restore the font size.
      ActiveChart.Legend.Font.Size = fntSZ
   End With

End Sub
Note: Make sure you have an Excel chart to run the code on the worksheet.
```

### **Uninstall Microsoft Work**

You may encounter a run-time error 1004 in Excel version 2009 or older versions due to conflicts between Microsoft works and Microsoft Excel. This error usually occurs if your system has both Microsoft Office and Microsoft Works. Uninstalling one of them will fix the issue. Try the below steps to uninstall Microsoft Work:

- First, open the **Task Manager** using the shortcut **CTRL + ALT + DEL** altogether
- The **Task Manager window** is displayed.

![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

### **Change Trust Center Settings**

Sometimes, run-time errors might arise because of incorrect security settings. The **Trust Center settings** help you find the **Privacy and security** settings for Microsoft Excel. Follow the below steps to change the **Trust center settings**:

- Open Microsoft Excel.
- Go to **File > Options.**
- The **Excel options** window is displayed.
- Choose **Trust Center**, and click **Trust Center Settings**.
- Tap on the **Macro Settings** tab, and select **Trust access to the VBA project object model.**

![Macro Settings in Microsoft Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/macro-settings.png)

- Click **OK**.

### **Run Open and Repair Tool**

The Runtime error also arises when MS Excel detects a corrupted worksheet. It automatically begins the File recovery mode and starts repairing it. However, if the Recovery mode fails to start, use the **Open and Repair** tool with the below steps:

- Click **File > Open**.
- Click the location and folder with a corrupted workbook.
- In the **Open** dialog box, choose the corrupted workbook.
- Click the arrow next to the **Open** tab, and go to the **Open and Repair** tab.
- Click **Repair**.

You can also opt for **Stellar Repair for Excel** if the Microsoft Excel’s built-in tool cannot fix the error.

### **Use Stellar Repair for Excel**

**Stellar Repair for Excel** is a professional software for repairing damage. xls, .xlsx, .xltm, .xltx, and .xlsm files and recovering all its objects. Here are the steps to fix the error using this tool:

- First, **download**, **install**, and run **Stellar Repair for Excel**.
- Click the **Browse** tab on the interface window to choose the corrupted Excel file you need to repair.
- Click **Scan**. You will see the scan progress in the scanning window.
- Click **OK**.
- The tool can let you preview all the recoverable Excel file components including tables, pivot tables, charts, formulas, etc.
- Click **Save** to save the repaired file.
- A **Save File dialog box** will appear with the below two options:
- Default location
- New location
- Choose a suitable option.
- Click the **Save** option to repair the Excel file that you have chosen.
- Once the repair is complete, it will display a message “**File repaired successfully**.”
- Click **OK**.

## **Conclusion**

Now you know the Excel run-time error 1004, its cause, and solutions. Follow the workarounds discussed in the blog to rectify the error quickly. However, **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** makes your task of removing run-time errors easy. It’s a powerful software to fix all the issues with Excel files. Also, it helps in extracting data from the damaged file and saves it to a new Excel workbook.




## How to Fix "Errors were detected while saving Excel" Error?

When trying to save the Excel file, you might face unexpected errors. The “Errors were detected while saving Excel” is one such error. It can also occur when using VBA in Excel. The complete error message appears as:

**“Errors were detected while saving \[file name\]. Microsoft Excel may be able to save the file by removing or repairing some features. To make the repairs in a new file, click Continue. To cancel saving the file, click Cancel.”  
**

The error can occur if the features (Pivot tables, charts, macros) used in the [Excel file get corrupted](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). However, there could be several other reasons behind the occurrence of the error. Let’s discuss them.

## **What Causes the "Errors were detected while saving Excel" Error?**

There are various reasons why you encounter this error. Here are some of them:

- Incompatible pivot table in the Excel file
- Large or uncompressed images in the Excel file
- File-sharing properties are not allowing file saving
- Excel file is corrupted
- Large-sized Excel file
- File version incompatibility
- VBA code is corrupted

## **Ways to Fix the “Errors were detected while saving Excel” Error**

You’re not able to save the Excel file if there is no storage space on your hard drive. So, first check if your hard drive has sufficient storage space to save the file. If this is not the case, then it might happen that your antivirus program is interrupting the saving process. To check this, temporarily disable your antivirus program and then try to save the file. If still your Excel is throwing the “Errors were detected while saving Excel” error, then follow the below given methods to fix the error:

###  **Method 1: Open the Excel in Safe Mode and Disable the Add-ins**

When you open Excel in safe mode, it opens without the third-party add-ins. This helps in finding out if any add-ins are causing the error.

 Here’s how to open the Excel in safe mode:

- Open the Run window by pressing **Windows key + R**.
- Type **excel /safe** in the Run window.  

    ![Excel Save Mode Command](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/excel-save-mode-command.jpg)?

- Next, click on **OK**.
- It will open Excel in safe mode.
- Now, try to open and save the affected file.

If you are able to save the file without any issue, then this means that the error has occurred due to third-party add-ins or settings. You can try disabling the add-ins to fix the issue. To do this, follow these steps:

- First, open Excel.
- Then, go to the **File** tab and click **Options**.

![Go To Options Window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/go-to-options-window.jpg)

- In **Excel Options**, click on the **Add-ins**

![Select Add-ins](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/select-add-ins4.jpg)

- Under the **Manage** section, select **Excel Add-ins** and then click on the **Go**

![Excel Add-ins Drop-down](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/excel-add-ins-drop-down.jpg)

- In the **Add-ins** dialog box, unselect the **add-ins** under the **Add-ins available** option and click **OK.**  

    ![Add-ins Window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/add-ins-window.jpg)

**_Note_**_: Disabling add-ins does not remove them from the system. To remove them permanently, you need to uninstall them._

### **Method 2: Check the Excel File Name**

Some users have observed this error when saving the Excel file with an invalid name. You can check the file name and ensure that it should not contain more than 218 characters. If the name exceeds the required limit, then try shortening the file name or move the file to a folder with a short path name.

###  **Method 3: Copy the Data from the Affected File to a New File**

If you are not able to save the Excel document, then try copying the data from the affected file to a new Excel file. Then, save the new file with a different name. This helps in resolving the issue.

### **Method 4: Check and Provide File Permissions**

You may experience the “Errors were detected while saving Excel" issue when you do not have desired permissions to modify the folder in which your Excel file is located. To modify the folder, you should have read, write, and create permissions. You can check and provide the desired permissions using the below steps:

- Navigate to the Windows **Program Files** and then find the desired folder (where the Excel file is saved).
- Right-click on the folder and then choose **Properties**.
- Select the **Security** tab and then click
- Click on **Change Permissions** in the **Advanced Settings**
- Click **Administrators** and then click **Edit**.
- Now set the **Apply to drop-down** button to **This Folder, Subfolder, and Files**.
- Click on the **Full Control** field and then click **Apply > OK**.

###  **Method 5: Check Pivot Tables in Excel Sheet**

You can review Pivot tables to see if they are causing the “Errors were detected while saving Excel” error. To do so, follow the below steps:

- Click **Power Pivot > Manage**.  

    ![Check Pivot Table In Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/check-pivot-table-in-excel.jpg)

- Check the tabs in the **Power Pivot**
- Check if all the formulas used in the table are correct. Sometimes, even a small typo can create an issue in Excel.

### **Method 6: Repair Your Excel File**

The “Errors were detected while saving Excel” issue can also occur if the Excel file is corrupted. In such a case, you can take the help of the built-in utility in Excel – Open and Repair to repair your Excel file. Here’s how to use the tool:

- In Excel, click the **File** tab and then click **Open**.
- Click **Browse** to select the desired file.
- The Open dialog box is displayed. Click on the corrupted file.
- Click on the arrow next to the **Open** button and then click **Open and Repair**.
- Click on the **Repair**

![Click On Repair Button](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/click-on-repair-button.jpg)

- After repair, a message will appear (as shown in the below figure).  

    ![Message Appear After Repair](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/message-appear-after-repair.jpg)

- Click **Close**.

 However, sometimes, the Open and Repair utility fails to fix the file if it is severely corrupted or large-sized. In such a case, you can take the help of a third-party Excel repair software, such as **[Stellar Repair for Excel.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** The tool performs a comprehensive scan of the corrupted Excel file to fix the issues and recover all the items from the file without changing the original formatting. It can recover pivot tables, charts, images, engineering formulas, etc. The tool is compatible with Windows 11/10/8.1/8/7. You can download the free trial version of the tool to evaluate its functionality.

##  **Closure**

Many Excel users reported facing the situation when they are saving the Excel file. You can check the file’s compatibility to fix the “Errors were detected while saving Excel” issue. If you are getting this error in a Macro-enabled file then you can try deleting the VBA project from a document to resolve the issue. However, deleting the entire VBA code cannot be a better solution as it can lead a data loss in the Project you are working on. In the above article, you have learned the reasons behind the issue and discovered how to fix the error. Follow the methods and if none of them works then try using Stellar Repair for Excel. It is an advanced tool that can quickly repair corruption in Excel worksheets at any level. It lets you restore the corrupted components from the corrupted file without removing the existing data.


## How to Resolve 'Excel found unreadable content in filename.xlsx' Error in MS Excel?

When opening an Excel spreadsheet in MS Office 2010/2007, you may get the following error message:

"Excel found unreadable content in '\[filename\].xlsx'. Do you want to recover the contents of this workbook? If you trust the source of this workbook, click Yes."

![Excel Found Unreadable Content Error Message](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Excel-found-unreadable-content-error-message-image-1.png)

On clicking 'Yes', you may face any of these scenarios:

**_Note:_** _If you choose to click 'No', then open your MS Excel application and click file > Open. When the Open dialog box opens, browse and select the file showing the 'Excel found unreadable content' error and then choose 'Open and Repair' option. If this didn't help, try using a third-party Excel repair tool to save time troubleshooting the issue and restoring the file with all its data intact._

**Scenario 1:** The following message may pop-up.

"Excel was able to open the file by repairing or removing the unreadable content. Excel recovered your formulas and cell values, but

[<u>some data may have been lost</u>](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

. Click to view log file listing repairs errorxxx.xml."

![Excel Was Able To Open the File By Repairing Message](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Excel-was-able-to-open-the-file-by-repairing-message-image-2.png)

The message clearly states that your Excel file might open, but images may be lost and other such inconsistencies can crop up.

**Scenario 2:** The error is followed by another error message, like "[<u>The file is corrupt and cannot be opened</u>](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)".

Watch our short video for a quick overview of the solutions to fix "Excel found unreadable content in filename.xlsx"

<iframe width="560" height="315" src="https://www.youtube.com/embed/6jYRjQAzwQ8?si=H4-22LK-s8Z3KwT9" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen=""></iframe>

## What Causes 'Excel Unreadable Content' Error?

You may encounter the 'Excel file unreadable content' error due to corruption of complete Excel file or corruption in certain areas (like Pivot Table, Formulas, Styles, or other objects) in the file. According to Microsoft, you may find it difficult to determine the root cause behind Excel file corruption. Corruption could occur in different scenarios, like power surge, a network glitch, copying and pasting corrupted data from another file, etc.

**Also Read**: [<u>How to recover data from&nbsp;corrupt or damaged&nbsp;Excel file 2010 &amp; 2007</u>](https://www.stellarinfo.com/article/recover-corrupted-excel-file-2010-2007.php)?

## Workarounds to Resolve the 'Excel found unreadable content in filename.xls' Error

There is no permanent solution to fix the 'Excel found unreadable content' error. But, following are some workarounds you can try to resolve the error.

**_Note:_** _Before you try any of these workarounds, run Excel with administrator privileges and try opening the Excel file that is throwing the 'unreadable content' error. If this doesn't fix the error, proceed with the workarounds below._

### **Workaround 1 – Try Opening the File in Excel 2003**

Sometimes a problem in the current Excel version might prevent a file from opening. To resolve this error, try opening the problematic file in Excel 2003. If the file opens, save the data in a web page file format (.html) and then try opening the .html file in MS Excel 2010/2007. The detailed step-wise instructions are as follows:

- Open the .xls file in Excel 2003.
- When the file opens, click on File > Save.
- In the 'Save As' dialog box, choose Web Page (.html) as the 'Save as type' and then click 'Save.' Doing so will save everything from your .xls file, opened with 2003, in .html file format.
- Open the .html file in Excel 2010/2007. And then, save the file with .xlsx extension with a new name to avoid overwriting the original file.

Now, open the Excel 2010/2007 file and check if the error is fixed. If not, use the next workaround.

### **Workaround 2 – Make the Excel File 'Read-only'**

Try to open your '.xlsx' file by making it 'read-only'. Follow these steps:

- In Excel, click 'File' from the main menu.
- Select 'Save' for a new document or 'Save As' for a previously saved document in the screen that appears.

![Excel File Saving Options](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Excel-file-saving-options-image-3.png)

- From the 'Save As' dialog box, click Tools > General Options.

![Open General Options In Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Open-general-options-in-excel-image-4.png)

- Click on the 'Read-only recommended' checkbox to make the document read-only and then click 'OK'.

![Select Read Only Recommended Option](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Select-read-only-recommended-option-image-5.png)

Now open a new '.xlsx' file and copy everything from the corrupt Excel file to this new file. Finally, save this file and try to open it again.

### **Workaround 3 – Move Excel File to a New Folder**

Some users have reported that they could open their Excel file, following the 'Excel unreadable content' error, by simply moving the file to a different folder and saving it under a new name. You can also move the affected file to a new folder and try opening it. If this didn't help resolve the error, follow the next workaround.

### **Workaround 4 – Install Visual Basic Component**

At times, it is seen that installing the 'Visual Basic' component of MS Office 2010 resolves the 'Excel found unreadable content 2010' error. To do so, follow these steps:

- Navigate to Control Panel > Programs and select Microsoft Office 2010.
- Click 'Change' and then select 'Add or Remove Programs'.
- Next, click the 'plus' sign provided next to Office Shared Features.
- Click 'Visual Basic for Applications'. After that, right-click and choose 'Run from My Computer' and hit the 'Continue' button.
- Reboot your system when this process finishes.

Now check if the issue has been resolved or not.

## What Next?

If none of the workarounds mentioned above works for you, use a professional [<u>Excel repair software</u>](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. The software repairs corrupt MS Excel sheets without modifying their original content and formatting. In addition, it can repair single or multiple Excel (XLS/XLSX) files in a few simple steps.

[![free-download](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/free-download-1-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Steps to Repair Corrupt Excel File using Stellar Repair for Excel Software

- Install and run Stellar Repair for Excel software.

- From the software main interface window, click 'Browse' to select the corrupt file. If you are not aware of the corrupt Excel file location, click on the 'Search' button.

![Select Corrupt excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Select-corrupt-excel-file-image-6-1024x544.png)

- Click on the 'Repair' button to scan and repair the selected file.

![Scan Corrupt Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Scan-corrupt-excel-file-image-7.png)

- A preview window will open with recoverable Excel file data. Once satisfied with the preview result, click on the 'Save File' button on the 'File' menu to start the repair process.

![Preview Recoverable Excel File Data](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Preview-recoverable-excel-file-data-image-8-1024x545.png)

- Select the destination to save the file.

![Save Repaired Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Save-repaired-excel-file-image-9.png)

- Click 'OK' when the 'Repaired file saved successfully' message appears.

![Saving Complete Message](https://www.stellarinfo.com/blog/wp-content/uploads/2021/07/Saving-complete-message-image-10.png)

 The repaired Excel file will get saved at the selected location.


## How to Repair Corrupted or Damaged Excel File with Ease?

**Summary:** The Excel file is prone to corruption. Users can face several issues related to corruption. So here in this infographic, I am discussing a professional tool,- Stellar Repair for Excel, to easily repair corrupted Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Stellar Repair for Excel is among the top choices for repairing corrupt or damaged Excel (.XLS/.XLSX) files. This [Excel recovery software](https://www.stellarinfo.com//blog/top-10-best-excel-recovery-software/) restores everything from the corrupt file to a new blank Excel file. Incoming, the information graphics complete overview of the repair process is explained in step-by-step methodology. Explore and reap the benefits of recovering corrupt or damaged Excel files.

[![Repair Corrupt Excel Files Infographic](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2024/02/Repair-Corrupt-Excel-Files-Infographic-2-scaled.jpg)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Very much sure about the result of using the excel file recovery tool, share your experience with us.



## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Many users have reported encountering the error “Subscript out of range” (runtime error 9) when using VBA code in Excel. The error often occurs when the object you are referring to in a code is not available, deleted, or not defined earlier. Sometimes, it occurs if you have declared an array in code but forgot to specify the DIM or ReDIM statement to define the length of array.

## **Causes of VBA Runtime Error 9: Subscript Out Of Range**

The error ‘Subscript out of range’ in Excel can occur due to several reasons, such as:

- Object you are trying to use in the VBA code is not defined earlier or is deleted.
- Entered a wrong declaration syntax of the array.
- Wrong spelling of the variable name.
- Referenced a wrong array element.
- Entered incorrect name of the worksheet you are trying to refer.
- Worksheet you trying to call in the code is not available.
- Specified an invalid element.
- Not specified the number of elements in an array.
- Workbook in which you trying to use VBA is corrupted.

## **Methods to Fix Excel VBA Error ‘Subscript out of Range’**

Following are some workarounds you can try to fix the runtime error 9 in Excel.

### **Method 1: Check the Name of Worksheet in the Code**

Sometimes, Excel throws the runtime error 9: Subscript out of range if the name of the worksheet is not defined correctly in the code. For example – When trying to copy content from one Excel sheet (emp) to another sheet (emp2) via VBA code, you have mistakenly mentioned wrong name of the worksheet (see the below code).

```
Private Sub CommandButton1_Click()
Worksheets("emp").Range("A1:E5").Select
Selection.Copy
Worksheets("emp3").Activate
Worksheets("emp3").Range("A1:E5").Select
ActiveSheet.Paste
Application.CutCopyMode = False
End Sub
```

![VBA Error Subscript Out Of Range-When Incorrect Name](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/VBA-error-subscript-out-of-range-when-incorrect-name.jpg)

When you run the above code, the Excel will throw the Subscript out of range error.

So, check the name of the worksheet and correct it. Here are the steps:

- Go to the **Design** tab in the **Developer** section.
- Double-click on the **Command** button.
- Check and modify the worksheet name (e.g. from “emp” to “emp2”).

![Modified Code From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/modified-code-from-emp-to-emp2.jpg)

- Now run the code.
- The content in ‘emp’ worksheet will be copied to ‘emp2’ (see below).

![Content Copied From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/content-copied-from-emp-to-emp2.jpg)

### Method 2: Check the Range of the Array

The VBA error “Subscript out of range” also occurs if you have declared an array in a code but didn’t specify the number of elements. For example – If you have declared an array and forgot to declare the array variable with elements, you will get the error (see below):

![Runtime Error 9 When Not Declared Array](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/run-time-error-9-when-not-declared-array.jpg)

To fix this, specify the array variable:

```
Sub FillArray()
Dim curExpense(364) As Currency
Dim intI As Integer
For intI = 0 to 364
curExpense(intI) = 20
Next
End Sub
```

### **Method 3: Change Macro Security Settings**

The Runtime error 9: Subscript out of range can also occur if there is an issue with the macros or macros are disabled in the Macro Security Settings. In such a case, you can check and change the macro settings. Follow these steps:

- Open your Microsoft Excel.
- Navigate to **File > Options > Trust Center**.
- Under **Trust Center**, select **Trust Center Settings**.
- Click **Macro Settings**, select **Enable all macros**, and then click **OK**.

![Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/macro-settings-in-trust-center.jpg)

### **Method 4: Repair your Excel File**

The name or format of the Excel file or name of the objects may get changed due to corruption in the file. When the objects are not identified in a VBA code, you may encounter the Subscript out of range error. You can use the Open and Repair utility in Excel to repair the corrupted file. To use this utility, follow these steps:

- In your MS Excel, click **File > Open**.
- Browse to the location where the affected file is stored.
- In the **Open** dialog box, select the corrupted workbook.
- In the **Open** dropdown, click on **Open and Repair**.
- You will see a prompt asking you to repair the file or extract data from it.
- Click on the **Repair** option to extract the data as much as possible. If **Repair** button fails, then click **Extract** button to recover data without formulas and values.

If the “Open and Repair” utility fails to repair the corrupted/damaged macro-enabled Excel file, then try an advanced Excel repair tool, such as Stellar Repair for Excel. It can easily repair severely corrupted Excel workbook and recover all the items, including macros, cell comments, table, charts, etc. with 100% integrity. The tool is compatible with all versions of Microsoft Excel.

## **Conclusion**

You may experience the “Subscript out of range” error while using VBA in Excel. You can follow the workarounds discussed in this blog to fix the error. If the Excel file is corrupt, then you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to repair the file. It’s a powerful software that can help fix all the issues that occur due to corruption in the Excel file. It helps to recover all the data from the corrupt Excel files (.xls, .xlsx, .xltm, .xltx, and .xlsm) without changing the original formatting. The tool supports Excel 2021, 2019, 2016, and older versions.


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
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-galaxy-z-fold-5-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Galaxy Z Fold 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-redmi-note-13-pro-5g-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Redmi Note 13 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-nokia-c110-by-fonelab-android-recover-music/"><u>How to restore wiped music on Nokia C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-motorola-g54-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Motorola G54 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-13-pro-without-a-computer-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 13 Pro without a computer?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-se-2020-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone SE (2020)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-xiaomi-redmi-note-13-proplus-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Xiaomi Redmi Note 13 Pro+ 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-erase-iphone-14-pro-max-data-permanently-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Erase iPhone 14 Pro Max Data Permanently | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-oppo-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-11-pro-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 11 Pro Data From iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-with-my-realme-narzo-60x-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to fix videos not playing with my Realme Narzo 60x 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-realme-c51-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Realme C51.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-realme-11x-5g-by-fonelab-android-recover-music/"><u>How to restore wiped music on Realme 11X 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-samsung-galaxy-s24plus-by-fonelab-android-recover-music/"><u>How to restore wiped music on Samsung Galaxy S24+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-honor-x7b-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Honor X7b Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-zte-blade-a73-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from ZTE Blade A73 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-samsung-galaxy-s21-fe-5g-2023-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Samsung Galaxy S21 FE 5G (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-oppo-a56s-5g-by-fonelab-android-recover-video/"><u>How to recover old videos from your Oppo A56s 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-xiaomi-redmi-13c-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Xiaomi Redmi 13C to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-realme-gt-5-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Realme GT 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-x-data-from-ios-itunes-backup-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone X Data From iOS iTunes Backup? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-samsung-galaxy-m54-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Samsung Galaxy M54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-meizu-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Meizu</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-honor-magic-6-pro-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Honor Magic 6 Pro PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-iphone-11-pro-from-backup-when-itunes-backup-is-corrupt-or-not-compatible-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore iPhone 11 Pro from Backup when iTunes Backup is Corrupt or not compatible | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-y100a-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Y100A on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-s17-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from S17</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-tecno-spark-20-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Tecno Spark 20?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-infinix-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Infinix</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-xiaomi-redmi-note-12-4g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Xiaomi Redmi Note 12 4G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-your-drivers-with-windows-device-manager-on-windows-11107-by-drivereasy-guide/"><u>How to identify some outdated your drivers with Windows Device Manager on Windows 11/10/7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-nubia-red-magic-9-pro-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Nubia Red Magic 9 Pro?</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-5-techniques-to-transfer-data-from-honor-x50-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 5 Techniques to Transfer Data from Honor X50 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/forgot-locked-apple-iphone-6-plus-password-learn-the-best-methods-to-unlock-by-drfone-ios/"><u>Forgot Locked Apple iPhone 6 Plus Password? Learn the Best Methods To Unlock</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-in-2024-create-like-a-pro-top-rated-animation-software-for-mac-and-pc/"><u>New In 2024, Create Like a Pro Top-Rated Animation Software for Mac and PC</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-within-the-realm-of-graphic-editing-people-also-ask-how-to-add-emoji-to-photo-stay-tuned-to-walk-through-the-best-possible-solutions-in-the-it-market/"><u>New Within the Realm of Graphic Editing, People Also Ask How to Add Emoji to Photo. Stay Tuned to Walk Through the Best Possible Solutions in the IT Market</u></a></li>
<li><a href="https://fake-location.techidaily.com/wondering-the-best-alternative-to-hola-on-lava-blaze-pro-5g-here-is-the-answer-drfone-by-drfone-virtual-android/"><u>Wondering the Best Alternative to Hola On Lava Blaze Pro 5G? Here Is the Answer | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/how-to-unlock-icloud-activation-lock-and-icloud-account-on-apple-iphone-xr-by-drfone-ios/"><u>How to Unlock iCloud Activation Lock and iCloud Account On Apple iPhone XR?</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-cast-honor-70-lite-5g-to-computer-for-iphone-and-android-drfone-by-drfone-android/"><u>How to Cast Honor 70 Lite 5G to Computer for iPhone and Android? | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/how-to-xiaomi-13t-pro-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Xiaomi 13T Pro Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://fix-guide.techidaily.com/motorola-edge-2023-screen-unresponsive-heres-how-to-fix-it-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Motorola Edge 2023 Screen Unresponsive? Heres How to Fix It | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-2024-approved-unleash-the-power-of-reverse-video-in-final-cut-pro/"><u>New 2024 Approved Unleash the Power of Reverse Video in Final Cut Pro</u></a></li>
<li><a href="https://review-topics.techidaily.com/how-to-upgrade-or-downgrade-iphone-se-2022-without-data-loss-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Upgrade or Downgrade iPhone SE (2022) Without Data Loss? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/android-safe-mode-how-to-turn-off-safe-mode-on-realme-narzo-60x-5g-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Android Safe Mode - How to Turn off Safe Mode on Realme Narzo 60x 5G? | Dr.fone</u></a></li>
</ul></div>



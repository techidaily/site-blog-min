---
title: How to Fix Excel 2021 Run Time Error 1004 | Stellar
date: 2024-05-19T18:32:11.593Z
updated: 2024-05-20T18:32:11.593Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2021 Run Time Error 1004
excerpt: This article describes How to Fix Excel 2021 Run Time Error 1004
keywords: repair damaged .xltx files,repair corrupt .xltx,repair excel 2003,repair corrupt .xlb,repair damaged excel file,repair excel
thumbnail: https://www.lifewire.com/thmb/TKoUz7zi8lw5cyOA93bOwDLfNYs=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/the_room-56cc7f225f9b5879cc590941.png
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


## How to repair 'recovered' Excel file that won't open

**Summary:** You may unable to open the Excel file after file recovery. When you try to open the file you may prompt up with error messages. This blog will discuss those errors and their solutions. Besides this, it also mentions Stellar Repair for Excel to repair and recover the severely corrupted Excel file with no data loss.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

**Consider a scenario:** A Windows 10 user accidentally deletes a folder containing multiple Excel (XLSX) files created in Excel 2013. The worst part was all the deleted files were important, and he did not have a backup, which would help him in restoring the data. Nonetheless, he was able to recover those data using a professional data recovery software. Next, he tried opening each of the recovered Excel files one by one. In doing so, his happiness turned into disappointment. And the reason was, some of the recovered Excel files failed to open prompting error messages, of course, due to corruption. In this scenario, all he needed was to repair ‘recovered’ Excel files that did not open. Know how to fix damaged Excel files that were recovered after deletion in this blog!

Before delving into the Excel file repair methods, acquaint yourself with the probable causes leading to damaged or corrupt Excel files. Although not essential, it shall be an added advantage in helping you to prevent Excel file corruption issues in the future.

## The standard Excel file(s) Corruption Causes

- **Power Outage** – A power outage due to which a system closes suddenly or unexpectedly is a common cause of damaged Excel files.
- **Forced System Shutdown** – Shutting down the system forcibly without closing MS Excel files is another common cause of corrupt Excel files.
- **PC Virus or Bug** – Computer viruses or bugs may affect one or more data file if owners do not protect their PCs with powerful updated antivirus. Same is the case with malware attacks. The chances are that your computer is infected with one of these and has affected a few Excel files.
- **Issues with Storage Devices** – Damaged or corrupted Excel files can also be the outcome of hard drive issues in the data storage systems. Logical hard drive damage is one of the most significant issues.

**The damaged Excel files may prompt up with error messages. These may be:**

- ‘Excel unable to read file’
- ‘Filename is not valid’
- ‘This file is not in a recognizable format’
- ‘abc.xls file cannot be accessed. The file may be read-only’
- ‘Excel found unreadable content in (filename)’, ‘The file is corrupt and cannot be opened’
- ‘Microsoft Excel has encountered a problem and needs to close’

**There exist multiple methods to repair Excel files which got damaged after recovery, and are as follows: XML method,**

- Open Excel files with HTML
- Inbuilt ‘Open and Repair’ feature/tool in Excel
- Repair damaged Excel file from TMP file by using Excel AutoRecover and AutoBackup features
- Recover data by configuring ‘calculation option’ as manual
- Moving/copying or transferring data to a different location
- Stellar Repair for Excel software

Taking Stellar software versus other methods, the former is a software solution that repairs damaged (XLS and XLSX) file quickly in major three steps: **Select**\->**Repair**\->**Save**, whereas others fall into the category of manual processes and consume both time and resources. Among the manual processes ‘**Open and Repair**’ method is the least time and resource consuming, which you can try quickly by following the below steps:

## Open and Repair method

1. Click **MS Excel** tab and then the **Open** tab
2. **Select** the damaged Excel file for repairing purpose
3. Click the arrow present beside the **Open** Next, click **Open and Repair** tab
4. Click either of the following:
    - **Repair** tab (recovers maximum data)
    - **Extract Data** tab (Recovers values and formulas if the repair process fails in recovering the complete data.)

**Note** – Use ‘Extract Data’ if ‘Repair’ is not successful.

If the Manual method is not competent enough to [repair recovered Excel files](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that don’t open then opt for a software-based solution of deploying **Stellar Repair for Excel software**.

<iframe title="How to Repair Excel File with Stellar Repair for Excel Software" width="750" height="422" frameborder="0" allowfullscreen="" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" nitro-og-src="https://www.youtube.com/embed/VAeGzHnETu0?feature=oembed&amp;autoplay=1" nitro-lazy-src="data:text/html;https://www.youtube.com/embed/VAeGzHnETu0?feature=oembed&amp;autoplay=1;base64,PGJvZHkgc3R5bGU9J3dpZHRoOjEwMCU7aGVpZ2h0OjEwMCU7bWFyZ2luOjA7cGFkZGluZzowO2JhY2tncm91bmQ6dXJsKGh0dHBzOi8vaW1nLnlvdXR1YmUuY29tL3ZpL1ZBZUd6SG5FVHUwLzAuanBnKSBjZW50ZXIvMTAwJSBuby1yZXBlYXQnPjxzdHlsZT5ib2R5ey0tYnRuQmFja2dyb3VuZDpyZ2JhKDAsMCwwLC42NSk7fWJvZHk6aG92ZXJ7LS1idG5CYWNrZ3JvdW5kOnJnYmEoMCwwLDApO2N1cnNvcjpwb2ludGVyO30jcGxheUJ0bntkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y2xlYXI6Ym90aDt3aWR0aDoxMDBweDtoZWlnaHQ6NzBweDtsaW5lLWhlaWdodDo3MHB4O2ZvbnQtc2l6ZTo0NXB4O2JhY2tncm91bmQ6dmFyKC0tYnRuQmFja2dyb3VuZCk7dGV4dC1hbGlnbjpjZW50ZXI7Y29sb3I6I2ZmZjtib3JkZXItcmFkaXVzOjE4cHg7dmVydGljYWwtYWxpZ246bWlkZGxlO3Bvc2l0aW9uOmFic29sdXRlO3RvcDo1MCU7bGVmdDo1MCU7bWFyZ2luLWxlZnQ6LTUwcHg7bWFyZ2luLXRvcDotMzVweH0jcGxheUFycm93e3dpZHRoOjA7aGVpZ2h0OjA7Ym9yZGVyLXRvcDoxNXB4IHNvbGlkIHRyYW5zcGFyZW50O2JvcmRlci1ib3R0b206MTVweCBzb2xpZCB0cmFuc3BhcmVudDtib3JkZXItbGVmdDoyNXB4IHNvbGlkICNmZmY7fTwvc3R5bGU+PGRpdiBpZD0ncGxheUJ0bic+PGRpdiBpZD0ncGxheUFycm93Jz48L2Rpdj48L2Rpdj48c2NyaXB0PmRvY3VtZW50LmJvZHkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBmdW5jdGlvbigpe3dpbmRvdy5wYXJlbnQucG9zdE1lc3NhZ2Uoe2FjdGlvbjogJ3BsYXlCdG5DbGlja2VkJ30sICcqJyk7fSk7PC9zY3JpcHQ+PC9ib2R5Pg=="></iframe>

Software with **100% integrity and precision**

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Conclusion

As an automated software, Stellar Repair for Excel is easy-to-use. You do not need technical know-how and skill set, as required to execute the manual processes. On analyzing all these methods, it is suggested to use Stellar Repair for Excel having a user-friendly GUI to repair Excel files that became damaged after recovery, or if the recovered Excel file(s) that you once deleted accidentally has become corrupt now. Above all, it is an excellent software with multiple features and advantages. Use it to address all your MS Excel issues or if any recovered Excel file is corrupt.




## Resolve Compile Error in Hidden Module in Excel: Causes & Solutions

The hidden module in Excel refers to a container with VBA codes, custom queries, and complex macros. The compile error in a hidden (protected) module in the Excel worksheet usually occurs when doing different activities on a macro-enabled sheet, such as merging .xls files. The error can result in macros execution failure. You need to quickly resolve this compile error to restore full functionality of the VBA code. Below, we’ll be discussing the solutions to fix this Excel error. But before that, let’s see why this error occurs.

You may encounter the Compile error in hidden module due to one of the following reasons:

- The code in the workbook is not compatible with the Excel application.
- Manual queries created in a previous version are no longer compatible with your current version of Excel.
- Missing references.
- Invalid .exe files (control information cache files) are automatically created with ActiveX control insertion in Excel file.
- Protected module is corrupted.
- The workbook with hidden module is damaged or corrupted.
- Incompatible add-ins.
- Incompatible Excel file version.
- The module is protected or password-protected.
- Missing or corrupted mscomctl.ocx file.

Excel can throw the compile error while compiling the code that exists in the protected module. So, first check the error and identify the hidden module that is creating the issue. You can unprotect the module. Also, ensure that you have permission to access the VBA code in the module. If the error still exists, follow the below troubleshooting methods.

### Method 1: Re-register ActiveX Control Files or mscomctl.ocx Files

You can get the compile error in the Excel file, containing the VBA code related to ActiveX controls or OCX files. The ActiveX control files and OCX files (mscomctl.ocx files) are the components of Microsoft’s standard controls library. The compile error in the hidden module can occur if these files are missing. In this case, you can use the Regsvr32 tool to re-register the OCX files. The [Regsvr32](https://support.microsoft.com/en-au/topic/how-to-use-the-regsvr32-tool-and-troubleshoot-regsvr32-error-messages-a98d960a-7392-e6fe-d90a-3f4e0cb543e5) is a command-line utility to register and unregister OLE controls in the Windows registry.

### Method 2: Delete .exd Files

 The .exd files are temporary files created by Excel when inserting ActiveX controls objects. These temporary files can lead to a compile error if they are corrupted. So, if this issue has occurred, particularly in the Excel file containing ActiveX controls, then deleting .exd files might fix the issue. To delete the .exd file, follow the below steps:

- First, open the **Run** window by pressing the Windows+R keys.

![Open The Run Window](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/open-the-run-window.jpg)

- In the **Run** window, type **%appdata%**.

![Type App Data Command](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/type-app-data-command.jpg)

- In the **Roaming** window, click on the **Microsoft** option.

![Click On Microsoft Option Under Roaming](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-microsoft-option-under-roaming.jpg)

- Under **Microsoft**, you will see a list of folders. Search and click on **Forms.**
- Right-click on a file with .exd extension and select **Delete**.
- Once you delete the .exd files, restart your Excel application.

### Method 3: Rollback the Office Updates

MS Office updates or upgrades may also cause the compile error in hidden module in Excel. If the error has occurred after downloading the recent Microsoft Office updates, try [reverting to the previous version](https://support.microsoft.com/en-us/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841) or uninstalling the recent updates to fix the issue.

### Method 4: Unselect Missing References

The compile error in hidden module determine path in Excel can also occur if your file contains a reference to object library/type library, which is labelled as Missing. You can locate, check, and uncheck the references marked as ‘Missing’ to fix the issue. Here are the steps:

- Open your **Excel** and press **Alt + F11** keys.
- The **Visual Basic Editor** is displayed.

![Visual Basic Editor](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/visual-basic-editor.jpg)

- Go to the **Tools** option and then click **References**.

![Click On References Under Tools Option](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-references-under-tools-option.jpg)

- In the **References-VBAProject** window, under **Available References**, search and unselect the references starting as “Missing”.

![Unselect Missing References](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/unselect-missing-references.jpg)

- Click **OK**.

### Method 5: Check the Code in Module

The compile error in hidden module can occur if there are issues in the code within the module. The problems include incorrect or missing syntaxes, missing parameters/references, or the code contains incompatible functions or a wrong name of the object. You can check and fix these issues in the code by opening the VBA editor.

### Method 6: Check and Remove Add-ins

In Excel, the compile error in macro-enabled files can also occur due to incompatible add-ins. You can check and disable the **add-ins** in Excel using the below steps:

- First, open the **Run** window and type excel /safe and then click **OK**. The Excel application will open in safe mode.
- Now try to open the affected Excel file. If it opens without the error, then check and remove the latest installed Excel add-ins.
- Navigate to the **File** option and then select **Options**.
- In the **Excel Options** window, click **Add-ins**.

![Click Addins Select Latest Addins](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-add-ins-select-latest-add-ins.jpg)

- Under **Add-ins**, search and select the latest add-ins, and then click on **Go**.
- In the **Add-ins** window, uncheck the add-ins and then click **OK**.

![Select  Analysis Toolpak](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/select-analysis-toolpak.jpg)

- Restart Excel and then check if the error is fixed or not.

### Method 7: Repair the Corrupt Excel File

Corruption in the Excel file can affect the macros in the hidden module, which may result in the compile error. In such a case, you can try repairing the Excel file using Microsoft’s inbuilt utility -Open and Repair. To use this tool, follow these steps:

- Open your Excel application.
- Click the **File** tab and then click **Open**.
- Click **Browse** to select the affected workbook.
- The **Open** dialog box will appear. Click on the corrupted file.
- Click the arrow next to the **Open** button and then **Open and Repair**.
- You will see a dialog box with three buttons - Repair, Extract Data, and Cancel.

![Click On Repair Option](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-repair-option-1.jpg)

- Click on the Repair button to recover as much of the data as possible.
- After repair, a message is displayed. Click **Close**.

![Message Appear After Repair](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/message-appear-after-repair.jpg)

## **What if None of the Above Solutions Works?**

If the above methods fail to get rid of the “compile error in hidden module” in Excel, then use an Excel repair tool such as Stellar Repair for Excel. This tool is specifically designed to repair the corrupted Excel file. It can recover all the components from corrupted Excel file (macros, queries, formulas, etc.) without changing their original formatting. The tool is compatible with all Excel versions and can be downloaded on a Windows system. You can download the free trial version of Stellar Repair for Excel to scan the corrupted Excel file and preview the data.

## **Closure**

You can get the “compile error in hidden module” when Excel detects any issue while compiling the code in a protected module. It can occur when there is an issue with the macro-enabled Excel workbook or Excel add-ins. You can follow the above-mentioned methods to fix the issue. If the error occurs due to corruption in the database file, then you can try [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It can repair severely corrupted Excel files. It also helps recover all the Excel workbook’s components, including macros and queries. The tool has a simple and user-friendly interface.


## \[Fixed\] Excel Cannot Open the File Because the File Format or File Extension Is Not Valid


When opening an older version of an Excel file in a newer version, you may encounter the "Excel cannot open the file because the extension is not valid" error. This happens if you have accidentally renamed the file with a different file format or an incorrect file extension. Also, you may get the 'Excel cannot open the file' error if the file has become unreadable or corrupted.

This is how the complete Excel error message looks like:

"Excel cannot open the file filename.xlsx because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file."

![Excel Cannot Open the File Extension Not Valid Error](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/excel-cannot-open-the-file-extension-not-valid-error-1.png)

  
Check out this video for a quick demonstration of how to fix "Excel Cannot Open the File Because the File Format or File Extension Is Not Valid"

<iframe width="560" height="315" src="https://www.youtube.com/embed/SobYKTdwY80?si=EOypjvXgpOQrgkqZ" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen=""></iframe>

## **How to Fix the 'Excel Cannot Open the File Because the File Format or Extension Is Not Valid' Error?**

Try any of these workarounds to fix the error and regain access to your Excel file data:

### **Workaround 1 – Change the Default File Format**

By default, an Excel file is saved in .xlsx or .xls file format. Trying to open the file with a different or incorrect extension may cause the 'Excel cannot open the file because the extension is not valid' error. To fix the error, try changing the Excel default file format by following these steps:

**_Note:_** _Take a backup of the Excel file that you cannot open and try the steps on the backup copy._

-   In Excel 2010 and newer versions, click on the **File** menu.

**_Note:_** _For Excel 2007 and earlier versions, click on the Office button._

-   On the left panel, click on **Options**. 
-   In 'Excel Options' window, under **Save workbooks** section, click on the **Save files in this file format:** drop-down. Select a file format (i.e., the one supported by your Excel version).

![Save Workbook in Other File Format](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/save-workbook-in-other-file-format-2.png)

-   Click **OK**.

Check if you can open your Excel file. If not, try the next workaround.

### **Workaround 2 – Edit the Excel File Permissions**

You may receive the 'Excel file format or file extension is not valid' error if you lack sufficient permissions to open the file. Usually this happens when you try to open an Excel file received from some other user. Follow these steps to change the file permissions and see if it fixes the problem:

-   Right-click on the file that won't open and click **Properties**.
-   In the file properties dialog box, click the **Security** tab and then hit the **Edit** button.

![Edit Excel File Properties](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/edit-excel-file-properties-3-1-e1624255637114.png)

-   When the file's permissions dialog box appears, click on the **Add** button.

![Add Excel File Permissions](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/add-excel-file-permissions-4.png)

-   Click on the **Advanced** button from the 'Select Users or Groups' window.

![Open Advanced Settings](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/open-advanced-settings-5.png)

-   Select **Find Now**. A list of all users and groups will get displayed in a search results box. Choose **Everyone** group from the list and hit **OK**.

![Select Everyone Group](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/select-everyone-group-6.png)

-   You can see the 'Everyone' group in the 'Enter the object names to select' textbox. Click **OK** once again to return to the file's permission dialog box.

![Everyone Group is Added Users Groups](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/everyone-group-is-added-users-groups-7.png)

-   Click **Everyone** from the Group or user names: box and check all the checkboxes under Allow.

![Allow Permissions for Everyone](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/allow-permissions-for-everyone-8-1-e1624260058974.png)

-   Click on **Apply** and then **OK**.

Now try to open the Excel file. If you're still getting the Excel file cannot open error, use the following workaround.

### **Workaround 3 – Recover Unsaved Workbook**

If the 'Excel file extension not valid error' occurs when attempting to open an unsaved workbook, do the following to recover the unsaved workbook:

-   In Excel, click **File.** 
-   From the Info screen, under **Manage Versions**, click **Recover Unsaved Workbooks**.

![Recover Unsaved Workbooks](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/recover-unsaved-workbooks-9.png)

Excel will list any unsaved files. Try opening the file and save it. If this doesn't work, skip to the next workaround.

### **Workaround 4 – Repair the Workbook**

If none of the above workarounds has worked for you, chances are that the Excel file has become corrupted. Try to repair the file using the Excel inbuilt 'Open and Repair' utility by following these steps:

-   In your Excel application, click on **File** > **Open**.
-   In the 'Open' dialog box that pops-up, select the Excel file you want to repair and click the arrow next to the **Open** button.

![Select Excel File for Repairing](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/select-excel-file-for-repairing-10.png)

-   From the dropdown list, select **Open and Repair**.

![Open and Repair Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/open-and-repair-excel-file-11.png)

-   Excel will ask you to attempt to repair the file or extract data from it. Click **Repair** to retrieve maximum data. If the Repair option fails, click on **Extract Data** to recover the data without formulas and values.

![Repair Excel File](https://www.stellarinfo.com/blog/wp-content/uploads/2021/06/repair-excel-file-12.jpg)

If the 'Open and Repair' utility doesn't help fix the corrupted file and/or recover the data, use an [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) such as Stellar Repair for Excel to regain access to your file and its data.  

[![free download](https://www.stellarinfo.com/blog/wp-content/uploads/2021/05/free-download-1-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Stellar Repair for Excel software can help repair severely corrupted XLS and XLSX files. Also, it helps recover all the file components, including tables, pivot tables, cell comments, charts, chart sheets, images, formulas, etc., without impacting the original structure of the Excel file. 

_**To learn how the software works, read this:**_ [**_How to repair corrupt Excel file using Stellar Repair for Excel?_**](https://www.stellarinfo.com/support/kb/index.php/article/repair-corrupt-excel-file)

## **Conclusion**

You may encounter the error 'Excel cannot open the file because the extension is not valid' when opening a workbook with a different file format or an incorrect file extension. Also, the error may occur if the Excel file has turned corrupt. This article has covered some of the most effective workarounds to resolve the error. But if nothing works, you can try to repair the Excel file and retrieve its data using the Stellar Repair for Excel software.


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
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-infinix-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Infinix ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-11-pro-max-to-other-iphone-15-pro-max-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 11 Pro Max To Other iPhone 15 Pro Max devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-motorola-razr-40-ultra-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Motorola Razr 40 Ultra Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h265-video-on-samsung-galaxy-xcover-6-pro-tactical-edition-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Samsung Galaxy XCover 6 Pro Tactical Edition?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-honor-magic-6-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Honor Magic 6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-x100-pro-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from X100 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-zte-blade-a73-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of ZTE Blade A73 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-6-ios-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 6 iOS? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-fodt-file-document-electronically-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign .fodt file document electronically</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-infinix-note-30-vip-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Infinix Note 30 VIP?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-yuva-2-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix corrupt video files of Yuva 2 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-xlsx-document-online-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to Sign a .xlsx document online</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-honor-play-8t-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Honor Play 8T Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-or-malfunctioning-your-drivers-with-windows-device-manager-on-windows-11-and-10-and-7-by-drivereasy-guide/"><u>How to identify missing or malfunctioning your drivers with Windows Device Manager on Windows 11 & 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-excel-2010-files-by-stellar-guide/"><u>How to Repair Corrupt Excel 2010 Files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-honor-magic-6-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Honor Magic 6.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-notes-from-iphone-se-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Notes from iPhone SE? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-max-to-other-iphone-11-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro Max To Other iPhone 11 devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-lava-yuva-2-by-fonelab-android-recover-music/"><u>How to restore wiped music on Lava Yuva 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-c51-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of C51 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-xltm-document-with-electronic-digital-signature-tutorial-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to Sign .xltm document with Electronic Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-dfu-mode-on-iphone-x-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of DFU Mode on iPhone X? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-drivers-with-windows-device-manager-in-windows-11-by-drivereasy-guide/"><u>How to identify malfunctioning drivers with Windows Device Manager in Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-vivo-x100-pro-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Vivo X100 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-itel-p55plus-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Itel P55+ to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-galaxy-s23-fe-get-deleted-phone-number-back-with-ease-and-safety-by-fonelab-android-recover-contacts/"><u>How to Galaxy S23 FE Get Deleted Phone Number Back with Ease and Safety</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-samsung-galaxy-xcover-7-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Samsung Galaxy XCover 7 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-vivo-t2-5g-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix corrupt video files of Vivo T2 5G using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-on-iphone-7-5-best-solutions-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover deleted pictures on iPhone 7? 5 Best Solutions | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-dead-iphone-14-pro-max-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to recover data from dead iPhone 14 Pro Max | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-iphone-xs-max-from-backup-when-itunes-backup-is-corrupt-or-not-compatible-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore iPhone XS Max from Backup when iTunes Backup is Corrupt or not compatible | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-smart-8-by-fonelab-android-recover-data/"><u>How to recover lost data from Smart 8?</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-realme-narzo-n53-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Realme Narzo N53 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-tecno-spark-10-pro-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Tecno Spark 10 Pro to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-14-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 14 System Issues? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-15-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 15 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-vivo-y36i-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Vivo Y36i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-motorola-edge-40-pro-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Motorola Edge 40 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-calendar-on-iphone-14-pro-max-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover lost Calendar on iPhone 14 Pro Max | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oneplus-nord-n30-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from OnePlus Nord N30 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-for-people-wanting-to-mock-gps-on-oneplus-ace-2v-devices-drfone-by-drfone-virtual/"><u>In 2024, For People Wanting to Mock GPS on OnePlus Ace 2V Devices | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-the-magnificent-art-of-pokemon-go-streaming-on-samsung-galaxy-a34-5g-drfone-by-drfone-virtual-android/"><u>In 2024, The Magnificent Art of Pokemon Go Streaming On Samsung Galaxy A34 5G? | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/edit-videos-without-watermark-restrictions-8-free-options/"><u>Edit Videos without Watermark Restrictions 8 Free Options</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-data-from-asus-rog-phone-7-ultimate-to-samsung-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Data from Asus ROG Phone 7 Ultimate to Samsung Phone | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/9-solutions-to-fix-poco-c51-system-crash-issue-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>9 Solutions to Fix Poco C51 System Crash Issue | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/how-to-check-if-your-samsung-galaxy-a34-5g-is-unlocked-by-drfone-android/"><u>How To Check if Your Samsung Galaxy A34 5G Is Unlocked</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/2024-approved-edit-like-a-pro-gopro-quik-review-and-best-pc-alternatives-for-video-editing/"><u>2024 Approved Edit Like a Pro GoPro Quik Review & Best PC Alternatives for Video Editing</u></a></li>
<li><a href="https://android-location-track.techidaily.com/how-to-intercept-text-messages-on-vivo-s17-pro-drfone-by-drfone-virtual-android/"><u>How to Intercept Text Messages on Vivo S17 Pro | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-erase-an-iphone-8-without-apple-id-by-drfone-ios/"><u>How to Erase an iPhone 8 without Apple ID?</u></a></li>
<li><a href="https://android-frp.techidaily.com/in-2024-is-gsm-flasher-adb-legit-full-review-to-bypass-your-nokia-130-music-phone-frp-lock-by-drfone-android/"><u>In 2024, Is GSM Flasher ADB Legit? Full Review To Bypass Your Nokia 130 Music Phone FRP Lock</u></a></li>
</ul></div>


